'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Licensee
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MKK       05/24/05                Original class definition
'  1.1       MR        06/05/05    Instantiated LicenseeCourse and LicenseeCourseTest
'                                   Objects and Added Functions for the same.
'  1.2       Manju     10/15/07     Added GetLicenseeStatus, GetLicenseeCertificationType
'            Hua Cao   09/08/12     Added GetManagerStatus
' Function          Description
' GetEntity(NAME)   Returns the Entity requested by the string arg NAME
' GetEntity(ID)     Returns the Entity requested by the int arg ID
' GetAll()          Returns an ReportsCollection with all Entity objects
' Add(ID)           Adds the Entity identified by arg ID to the 
'                           internal ReportsCollection
' Add(Name)         Adds the Entity identified by arg NAME to the internal 
'                           ReportsCollection
' Add(Entity)       Adds the Entity passed as the argument to the internal 
'                           ReportsCollection
' Remove(ID)        Removes the Entity identified by arg ID from the internal 
'                           ReportsCollection
' Remove(NAME)      Removes the Entity identified by arg NAME from the 
'                           internal ReportsCollection
' EntityTable()     Returns a datatable containing all columns for the Entity 
'                          objects in the internal ReportsCollection.
''-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLicensee
#Region "Public Events"
        Public Event LicenseeErr(ByVal MsgStr As String)
        Public Event LicenseeChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
        Public Event GenerateCongLetter()
        Public Event GenerateLicenseeCertLetter()
        Public Event GenerateLicenseeCard()
        Public Event LicenseeCourseChanged(ByVal bolValue As Boolean)
        Public Event LicenseeCourseTestsChanged(ByVal bolValue As Boolean)
        Public Event GenerateNoCertificationLetter(ByVal oLicensee As MUSTER.Info.LicenseeInfo)
        Public Event GenerateLicenseeRenewalLetter()
        Public Event GenerateNoCertificationLetterOption(ByVal ByValoLicensee As MUSTER.Info.LicenseeInfo)
#End Region
#Region "Private Member Variables"
        Private WithEvents oLicenseeInfo As MUSTER.Info.LicenseeInfo
        Private WithEvents colLicensees As MUSTER.Info.LicenseeCollection
        Private WithEvents colManagers As MUSTER.Info.LicenseeCollection
        Private oLicenseeDB As New MUSTER.DataAccess.LicenseeDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private nCompAddID As Integer = 0
        Private colKey As String = String.Empty
        Private strStatus As String = String.Empty
        Private WithEvents pLicCourse As MUSTER.BusinessLogic.pLicenseeCourses
        Private WithEvents pMgrFacRelation As MUSTER.BusinessLogic.pManagerFacRelations
        Private WithEvents pLicCourseTest As MUSTER.BusinessLogic.pLicenseeCourseTest
        Private WithEvents oComments As MUSTER.BusinessLogic.pComments
#End Region
#Region "Constructors"
        Public Sub New()
            oLicenseeInfo = New MUSTER.Info.LicenseeInfo
            colLicensees = New MUSTER.Info.LicenseeCollection
            colManagers = New MUSTER.Info.LicenseeCollection
            pLicCourse = New MUSTER.BusinessLogic.pLicenseeCourses
            pMgrFacRelation = New MUSTER.BusinessLogic.pManagerFacRelations
            pLicCourseTest = New MUSTER.BusinessLogic.pLicenseeCourseTest
            oComments = New MUSTER.BusinessLogic.pComments
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property LicenseeInfo() As MUSTER.info.LicenseeInfo
            Get
                Return oLicenseeInfo
            End Get
        End Property
        Public ReadOnly Property Licensee_name() As String
            Get
                Return oLicenseeInfo.TITLE + " " + oLicenseeInfo.FIRST_NAME + " " + oLicenseeInfo.MIDDLE_NAME + " " + oLicenseeInfo.LAST_NAME + " " + oLicenseeInfo.SUFFIX
            End Get
        End Property
        Public Property colLicensee() As MUSTER.Info.LicenseeCollection
            Get
                Return colLicensees
            End Get
            Set(ByVal Value As MUSTER.Info.LicenseeCollection)
                colLicensees = Value
            End Set
        End Property
        Public Property colManager() As MUSTER.Info.LicenseeCollection
            Get
                Return colManagers
            End Get
            Set(ByVal Value As MUSTER.Info.LicenseeCollection)
                colManagers = Value
            End Set
        End Property
        Public Property COMP_ADD_ID() As Integer
            Get
                Return nCompAddID
            End Get
            Set(ByVal Value As Integer)
                nCompAddID = Value
            End Set
        End Property
        Public Property APP_RECVD_DATE() As DateTime
            Get
                Return oLicenseeInfo.APP_RECVD_DATE
            End Get
            Set(ByVal Value As DateTime)
                oLicenseeInfo.APP_RECVD_DATE = Value
            End Set
        End Property
        Public Property TITLE() As String
            Get
                Return oLicenseeInfo.TITLE
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.TITLE = Value
            End Set
        End Property
        Public Property ASSOCATED_COMPANY_ID() As Integer
            Get
                Return oLicenseeInfo.ASSOCATED_COMPANY_ID
            End Get
            Set(ByVal Value As Integer)
                oLicenseeInfo.ASSOCATED_COMPANY_ID = Value
            End Set
        End Property
        Public Property CERT_TYPE_ID() As String
            Get
                Return oLicenseeInfo.CertType
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.CertType = Value
            End Set
        End Property
        Public Property CMCERT_TYPE_ID() As String
            Get
                Return oLicenseeInfo.CMCertType
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.CMCertType = Value
            End Set
        End Property
        Public Property CERT_TYPE_DESC() As String
            Get
                Return oLicenseeInfo.CertTypeDesc
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.CertTypeDesc = Value
            End Set
        End Property
        Public Property CMCERT_TYPE_DESC() As String
            Get
                Return oLicenseeInfo.CMCertTypeDesc
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.CMCertTypeDesc = Value
            End Set
        End Property
        Public Property DELETED() As Boolean
            Get
                Return oLicenseeInfo.DELETED
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeInfo.DELETED = Value
            End Set
        End Property
        Public Property EMAIL_ADDRESS() As String
            Get
                Return oLicenseeInfo.EMAIL_ADDRESS
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.EMAIL_ADDRESS = Value
            End Set
        End Property
        Public Property EMPLOYEE_LETTER() As Boolean
            Get
                Return oLicenseeInfo.EMPLOYEE_LETTER
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeInfo.EMPLOYEE_LETTER = Value
            End Set
        End Property
        Public Property EXCEPT_GRANT_DATE() As DateTime
            Get
                Return oLicenseeInfo.EXCEPT_GRANT_DATE
            End Get
            Set(ByVal Value As DateTime)
                oLicenseeInfo.EXCEPT_GRANT_DATE = Value
            End Set
        End Property
        Public Property FIRST_NAME() As String
            Get
                Return oLicenseeInfo.FIRST_NAME
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.FIRST_NAME = Value
            End Set
        End Property
        Public Property HIRE_STATUS() As String
            Get
                Return oLicenseeInfo.HIRE_STATUS
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.HIRE_STATUS = Value
            End Set
        End Property
        Public Property ID() As Int64
            Get
                Return oLicenseeInfo.ID
            End Get
            Set(ByVal Value As Int64)
                oLicenseeInfo.ID = Value
            End Set
        End Property
        Public Property ISSUED_DATE() As Date
            Get
                Return oLicenseeInfo.ISSUED_DATE
            End Get
            Set(ByVal Value As Date)
                oLicenseeInfo.ISSUED_DATE = Value
            End Set
        End Property
        Public Property EXTENSION_DEADLINE_DATE() As Date
            Get
                Return oLicenseeInfo.EXTENSION_DEADLINE_DATE
            End Get
            Set(ByVal Value As Date)
                oLicenseeInfo.EXTENSION_DEADLINE_DATE = Value
            End Set
        End Property
        Public Property COMPLIANCEMANAGER() As Boolean
            Get
                Return oLicenseeInfo.COMPLIANCEMANAGER
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeInfo.COMPLIANCEMANAGER = Value
            End Set
        End Property
        Public Property ISLICENSEE() As Boolean
            Get
                Return oLicenseeInfo.ISLICENSEE
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeInfo.ISLICENSEE = Value
            End Set
        End Property
        Public Property INITCERTDATE() As String
            Get
                Return oLicenseeInfo.INITCERTDATE
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.INITCERTDATE = Value
            End Set
        End Property
        Public Property INITCERTBY() As Integer
            Get
                Return oLicenseeInfo.INITCERTBY
            End Get
            Set(ByVal Value As Integer)
                oLicenseeInfo.INITCERTBY = Value
            End Set
        End Property
        Public Property INITCERTBYDESC() As String
            Get
                Return oLicenseeInfo.INITCERTBYDESC
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.INITCERTBYDESC = Value
            End Set
        End Property
        Public Property RETRAINDATE1() As String
            Get
                Return oLicenseeInfo.RETRAINDATE1
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.RETRAINDATE1 = Value

            End Set
        End Property
        Public Property RETRAINDATE2() As String
            Get
                Return oLicenseeInfo.RETRAINDATE2
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.RETRAINDATE2 = Value

            End Set
        End Property
        Public Property RETRAINDATE3() As String
            Get
                Return oLicenseeInfo.RETRAINDATE3
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.RETRAINDATE3 = Value

            End Set
        End Property
        Public Property REVOKEDATE() As String
            Get
                Return oLicenseeInfo.REVOKEDATE
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.REVOKEDATE = Value

            End Set
        End Property
        Public Property RETRAINREQDATE() As String
            Get
                Return oLicenseeInfo.RETRAINREQDATE
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.RETRAINREQDATE = Value

            End Set
        End Property
        Public Property LAST_NAME() As String
            Get
                Return oLicenseeInfo.LAST_NAME
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.LAST_NAME = Value
            End Set
        End Property
        Public Property LICENSE_EXPIRE_DATE() As String
            Get
                Return oLicenseeInfo.LICENSE_EXPIRE_DATE
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.LICENSE_EXPIRE_DATE = Value
            End Set
        End Property
        Public Property MIDDLE_NAME() As String
            Get
                Return oLicenseeInfo.MIDDLE_NAME
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.MIDDLE_NAME = Value
            End Set
        End Property
        Public ReadOnly Property FullName() As String
            Get
                Return oLicenseeInfo.FULLNAME
            End Get
        End Property
        Public Property ORGIN_ISSUED_DATE() As Date
            Get
                Return oLicenseeInfo.ORIGIN_ISSUED_DATE
            End Get
            Set(ByVal Value As Date)
                oLicenseeInfo.ORIGIN_ISSUED_DATE = Value
            End Set
        End Property
        Public Property OVERRIDE_EXPIRE() As Boolean
            Get
                Return oLicenseeInfo.OVERRIDE_EXPIRE
            End Get
            Set(ByVal Value As Boolean)
                oLicenseeInfo.OVERRIDE_EXPIRE = Value
            End Set
        End Property
        Public Property STATUS_ID() As Integer
            Get
                Return oLicenseeInfo.STATUS
            End Get
            Set(ByVal Value As Integer)
                oLicenseeInfo.STATUS = Value
            End Set
        End Property
        Public Property STATUS_DESC() As String
            Get
                Return oLicenseeInfo.StatusDesc
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.StatusDesc = Value
            End Set
        End Property
        Public Property CMSTATUS_ID() As Integer
            Get
                Return oLicenseeInfo.CMSTATUS
            End Get
            Set(ByVal Value As Integer)
                oLicenseeInfo.CMSTATUS = Value
            End Set
        End Property
        Public Property CMSTATUS_DESC() As String
            Get
                Return oLicenseeInfo.CMStatusDesc
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.CMStatusDesc = Value
            End Set
        End Property
        Public Property SUFFIX() As String
            Get
                Return oLicenseeInfo.SUFFIX
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.SUFFIX = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLicenseeInfo.IsDirty
            End Get
            Set(ByVal value As Boolean)
                oLicenseeInfo.IsDirty = Boolean.Parse(value)
                RaiseEvent LicenseeChanged(value)
            End Set
        End Property
        Public ReadOnly Property colIsDirty() As Boolean
            Get
                If oLicenseeInfo.IsDirty Or pLicCourse.colIsDirty Or pLicCourseTest.colIsDirty Then
                    Return True
                    Exit Property
                End If
                Return False
            End Get
        End Property

        Public Property pLicenseeCourse() As MUSTER.BusinessLogic.pLicenseeCourses
            Get
                Return pLicCourse
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pLicenseeCourses)
                pLicCourse = Value
            End Set
        End Property
        Public Property pManagerFacRelation() As MUSTER.BusinessLogic.pManagerFacRelations
            Get
                Return pMgrFacRelation
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pManagerFacRelations)
                pMgrFacRelation = Value
            End Set
        End Property
        Public Property pLicenseeCourseTest() As MUSTER.BusinessLogic.pLicenseeCourseTest
            Get
                Return pLicCourseTest
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pLicenseeCourseTest)
                pLicCourseTest = Value
            End Set
        End Property
        Public Property Comments() As MUSTER.BusinessLogic.pComments
            Get
                Return oComments
            End Get
            Set(ByVal Value As MUSTER.BusinessLogic.pComments)
                oComments = Value
            End Set
        End Property
        Public Property LICENSEE_NUMBER() As String
            Get
                Return oLicenseeInfo.LICENSEE_NUMBER
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.LICENSEE_NUMBER = Value
            End Set
        End Property
        Public Property LICENSEE_NUMBER_PREFIX() As String
            Get
                Return oLicenseeInfo.LICENSEE_NUMBER_PREFIX
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.LICENSEE_NUMBER_PREFIX = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oLicenseeInfo.CREATED_BY
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.CREATED_BY = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLicenseeInfo.DATE_CREATED
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oLicenseeInfo.LAST_EDITED_BY
            End Get
            Set(ByVal Value As String)
                oLicenseeInfo.LAST_EDITED_BY = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLicenseeInfo.DATE_LAST_EDITED
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.LicenseeInfo
            Dim oLicenseeInfoLocal As MUSTER.Info.LicenseeInfo
            Try
                oLicenseeInfo = colLicensees.Item(ID)
                If Not oLicenseeInfo Is Nothing Then
                    Return oLicenseeInfo
                End If
                oLicenseeInfo = oLicenseeDB.DBGetByID(ID)
                If oLicenseeInfo.ID = 0 Then
                    oLicenseeInfo.ID = nID
                    nID -= 1
                End If
                colLicensees.Add(oLicenseeInfo)
                Return oLicenseeInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False) As Boolean
            Try
                'If Not bolValidated Then
                '    'If Not Me.ValidateData() Then
                '    '    Return False
                '    'End If
                'End If

                If Not ((oLicenseeInfo.ID < 0 And oLicenseeInfo.ID > -100) And oLicenseeInfo.DELETED) Then
                    Dim OldKey As String = oLicenseeInfo.ID.ToString
                    oLicenseeDB.Put(oLicenseeInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    If Not bolValidated Then
                        If oLicenseeInfo.ID.ToString <> OldKey Then
                            colLicensees.ChangeKey(OldKey, oLicenseeInfo.ID.ToString)
                        End If
                    End If
                    oLicenseeInfo.Archive()
                    oLicenseeInfo.IsDirty = False
                End If

                'Save all Licensee Courses and LicenseeCourseTests
                pLicenseeCourse.Flush(moduleID, staffID, returnVal, oLicenseeInfo.ID)
                If Not returnVal = String.Empty Then
                    Exit Function
                End If
                pLicenseeCourseTest.Flush(moduleID, staffID, returnVal, oLicenseeInfo.ID)
                If Not returnVal = String.Empty Then
                    Exit Function
                End If
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function CMSave(ByVal moduleID As Integer, ByVal staffID As Integer, ByVal CompanyAddrID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal CompanyID As Integer = 0) As Boolean
            Try
                'If Not bolValidated Then
                '    'If Not Me.ValidateData() Then
                '    '    Return False
                '    'End If
                'End If

                If Not ((oLicenseeInfo.ID < 0 And oLicenseeInfo.ID > -100) And oLicenseeInfo.DELETED) Then
                    Dim OldKey As String = oLicenseeInfo.ID.ToString
                    oLicenseeDB.CMPut(oLicenseeInfo, moduleID, staffID, returnVal, CompanyID, CompanyAddrID)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    If Not bolValidated Then
                        If oLicenseeInfo.ID.ToString <> OldKey Then
                            colLicensees.ChangeKey(OldKey, oLicenseeInfo.ID.ToString)
                        End If
                    End If
                    oLicenseeInfo.Archive()
                    oLicenseeInfo.IsDirty = False
                End If

                'Save all Licensee Courses and LicenseeCourseTests
                '  pLicenseeCourse.Flush(moduleID, staffID, returnVal, oLicenseeInfo.ID)
                pManagerFacRelation.Flush(moduleID, staffID, returnVal, oLicenseeInfo.ID)
                If Not returnVal = String.Empty Then
                    Exit Function
                End If
                'pLicenseeCourseTest.Flush(moduleID, staffID, returnVal, oLicenseeInfo.ID)
                'If Not returnVal = String.Empty Then
                '    Exit Function
                'End If
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        'Public Sub GenerateLicenseeLogic(ByRef oLicensee As MUSTER.Info.LicenseeInfo, ByVal FinRespExpDate As Date, Optional ByVal bolCompany As Boolean = False, Optional ByVal LicID As Integer = 0)
        '    Dim bolFinanciallyResponsible As Boolean
        '    Dim bolInstall As Boolean = False
        '    Dim bolClosure As Boolean = False
        '    Dim bolInstall1 As Boolean = False
        '    Dim bolClosure1 As Boolean = False
        '    Dim xLicCourse As MUSTER.Info.LicenseeCourseInfo
        '    Dim bolCourseDate As Boolean = False

        '    Try
        '        If bolCompany Then
        '            If oLicensee.HIRE_STATUS <> "RX - Not for Hire - Owner" Then
        '                If Date.Compare(FinRespExpDate.Date, Now.Date) > 0 Then
        '                    If oLicensee.EMPLOYEE_LETTER Or oLicensee.HIRE_STATUS = "HX - For Hire - Owner" Then
        '                        bolFinanciallyResponsible = True
        '                    Else
        '                        bolFinanciallyResponsible = False
        '                    End If
        '                Else
        '                    bolFinanciallyResponsible = False
        '                End If
        '            Else
        '                bolFinanciallyResponsible = True
        '            End If
        '        Else
        '            bolFinanciallyResponsible = False
        '        End If
        '        'Licensee contains a closure test and an install test with each having a score of alteast 75 and a test date > today - 6    months()
        '        Dim xlicCourseTestInfo As MUSTER.Info.LicenseeCourseTestInfo
        '        For Each xlicCourseTestInfo In pLicCourseTest.colLicCourseTest.Values
        '            If oLicensee.StatusDesc <> "CERTIFIED" Then
        '                If (xlicCourseTestInfo.CourseTypeID = 920 And xlicCourseTestInfo.TestScore > 74 And DateDiff(DateInterval.Month, xlicCourseTestInfo.TestDate, Now.Date) < 6) Then
        '                    bolInstall = True
        '                End If
        '                If (xlicCourseTestInfo.CourseTypeID = 921 And xlicCourseTestInfo.TestScore > 74 And DateDiff(DateInterval.Month, xlicCourseTestInfo.TestDate, Now.Date) < 6) Then
        '                    bolClosure = True
        '                End If
        '            ElseIf oLicensee.StatusDesc = "CERTIFIED" Then
        '                If (xlicCourseTestInfo.CourseTypeID = 920 And xlicCourseTestInfo.TestScore > 74 And DateDiff(DateInterval.Year, xlicCourseTestInfo.TestDate, Now.Date) <= 2) Then
        '                    bolInstall = True
        '                End If
        '                If (xlicCourseTestInfo.CourseTypeID = 921 And xlicCourseTestInfo.TestScore > 74 And DateDiff(DateInterval.Year, xlicCourseTestInfo.TestDate, Now.Date) <= 2) Then
        '                    bolClosure = True
        '                End If
        '            End If


        '        Next
        '        ' Licensee contains a closure course and an install                 course()
        '        Dim xlicCourseInfo As MUSTER.Info.LicenseeCourseInfo
        '        For Each xlicCourseInfo In pLicCourse.colLicCourse.Values()
        '            If xlicCourseInfo.CourseTypeID = 920 Then
        '                bolInstall1 = True
        '            End If
        '            If xlicCourseInfo.CourseTypeID = 921 Then
        '                bolClosure1 = True
        '            End If
        '        Next
        '        If bolClosure And bolInstall Then
        '            If bolClosure1 And bolInstall1 Then
        '                oLicenseeInfo.CertTypeDesc = "INSTALL"
        '            ElseIf bolClosure1 Then
        '                oLicenseeInfo.CertTypeDesc = "CLOSURE"
        '            End If
        '        ElseIf bolClosure And bolClosure1 Then
        '            oLicenseeInfo.CertTypeDesc = "CLOSURE"
        '        Else
        '            oLicenseeInfo.CertTypeDesc = "NONE"
        '            oLicensee.StatusDesc = "APPLICANT"
        '        End If
        '        If oLicenseeInfo.CertTypeDesc = "INSTALL" Or oLicenseeInfo.CertTypeDesc = "CLOSURE" Then
        '            If bolFinanciallyResponsible Then
        '                'replace code with - if date is not null
        '                If Date.Compare(oLicensee.APP_RECVD_DATE, CDate("01/01/0001")) <> 0 Then
        '                    oLicensee.StatusDesc = "CERTIFIED"
        '                Else
        '                    oLicensee.StatusDesc = "APPLICANT"
        '                End If
        '            Else
        '                oLicensee.StatusDesc = "APPLICANT"
        '            End If
        '        End If

        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        ''Executes the licensee logic
        'Public Sub LicenseeLogic(ByRef oLicensee As MUSTER.Info.LicenseeInfo, ByVal FinRespExpDate As Date, Optional ByVal bolCompany As Boolean = False, Optional ByVal LicID As Integer = 0, Optional ByVal strLicenseeHireStatus As String = "")
        '    Dim xLicCourse As MUSTER.Info.LicenseeCourseInfo
        '    Dim bolCourseDate As Boolean = False
        '    Dim mostRecentCourseDate As DateTime = CDate("01/01/2001")
        '    Dim priorStatus As String = oLicensee.StatusDesc
        '    Dim priorCertType As String = oLicenseeInfo.CertTypeDesc
        '    Dim bolRenewing As Boolean = False

        '    Try
        '        oLicenseeInfo = oLicensee
        '        If oLicensee.StatusDesc = "" Or oLicensee.StatusDesc = "APPLICANT" Or oLicensee.StatusDesc = "NO LONGER WITH COMPANY" Then
        '            Me.GenerateLicenseeLogic(oLicensee, FinRespExpDate, bolCompany, LicID)  '-- applicant logic

        '        ElseIf oLicensee.StatusDesc = "NOT CURRENTLY CERTIFIED" Then
        '            If oLicensee.OVERRIDE_EXPIRE Then
        '                Me.renewalLicensee(oLicensee, FinRespExpDate, bolCompany)
        '                bolRenewing = True
        '            Else
        '                Me.expiredlicensee(oLicensee, FinRespExpDate, bolCompany)
        '            End If
        '        ElseIf Date.Compare(Now.Date, oLicensee.LICENSE_EXPIRE_DATE) > 0 Then
        '            Me.expiredlicensee(oLicensee, FinRespExpDate, bolCompany)

        '            '----- check to see if the status should be changed from CERTIFIED to RENEWING
        '        ElseIf oLicensee.StatusDesc = "CERTIFIED" And _
        '                Date.Compare(oLicensee.APP_RECVD_DATE, oLicensee.ISSUED_DATE) > 0 And _
        '                DateDiff(DateInterval.Day, Now, oLicensee.LICENSE_EXPIRE_DATE) < 60 Then
        '            Me.renewalLicensee(oLicensee, FinRespExpDate, bolCompany)
        '            bolRenewing = True

        '            '----- if the status is RENEWING, then run the RENEWING logic until it becomes CERTIFIED 
        '        ElseIf oLicensee.StatusDesc = "RENEWING" Then
        '            Me.renewalLicensee(oLicensee, FinRespExpDate, bolCompany)
        '            bolRenewing = True

        '            '----- if the status is CERTIFIED then run the Applicant logic 
        '        ElseIf oLicensee.StatusDesc = "CERTIFIED" Then
        '            Me.GenerateLicenseeLogic(oLicensee, FinRespExpDate, bolCompany, LicID)   '-- applicant process 
        '        End If

        '        '-------- effective end of status examination and changing logic ... 
        '        '        code below is what will happen after the status changing logic has executed 
        '        '        (print letters, etc..)
        '        '======================================================================================================
        '        If oLicensee.StatusDesc = "CERTIFIED" And LicID = 0 Then
        '            'Generate the congratulatory letter
        '            'Generate the License Certificate corresponding to the certification type
        '            'Generate the Licensee Cardv
        '            RaiseEvent GenerateCongLetter()
        '            RaiseEvent GenerateLicenseeCertLetter()
        '            RaiseEvent GenerateLicenseeCard()
        '            '-- set Original Issue Date and License Issue Date as Today
        '            '-- set License Expiration Date as first day of the next month plus 2 years
        '            oLicensee.ISSUED_DATE = Now.Date
        '            oLicensee.ORIGIN_ISSUED_DATE = Now.Date
        '            Dim expression As String = IIf(Now.Month = 12, "01", (Now.Month + 1).ToString) + "/" + "01/" + (Now.Year + 2).ToString
        '            oLicensee.LICENSE_EXPIRE_DATE = CDate(expression)
        '            'Update the company's Type
        '        End If

        '        '-----------------------------------------------------------------------------------------------
        '        'For modify execute the following logic
        '        If LicID <> 0 Then
        '            '----- if the License Issued Date and the License Expiration Date is not null 
        '            If (Date.Compare(oLicensee.ISSUED_DATE, CDate("01/01/0001")) > 0 And _
        '                Date.Compare(oLicensee.LICENSE_EXPIRE_DATE, CDate("01/01/0001")) > 0) Then
        '                '----- if the Application Received Date is Greater than the License Issued Date 
        '                '       AND the License Expiration Date is within 60 days of today 
        '                '       AND the status is CERTIFIED 
        '                '       This reflects a licensee that has come from Renewal to Certified ------ 
        '                If Date.Compare(oLicensee.APP_RECVD_DATE, oLicensee.ISSUED_DATE) > 0 And _
        '                    (DateDiff(DateInterval.Day, Now, oLicensee.LICENSE_EXPIRE_DATE) < 60) And _
        '                    (oLicensee.StatusDesc = "CERTIFIED") And _
        '                    (bolRenewing = True) Then
        '                    '----- this indicates that it went through the renewing process and came out 
        '                    '      certified 
        '                    '----- check all the courses and find if there is a course date greater than expiration date --
        '                    For Each xLicCourse In pLicCourse.colLicCourse.Values
        '                        If Date.Compare(xLicCourse.CourseDate, oLicensee.LICENSE_EXPIRE_DATE) > 0 Then 'xLicCourse.CourseDate > oLicensee.LICENSE_EXPIRE_DATE Then
        '                            bolCourseDate = True
        '                        End If
        '                        If Date.Compare(mostRecentCourseDate, xLicCourse.CourseDate) < 0 Then 'mostRecentCourseDate < xLicCourse.CourseDate Then
        '                            mostRecentCourseDate = xLicCourse.CourseDate
        '                        End If
        '                    Next

        '                    If bolCourseDate Then
        '                        Dim msgResult As MsgBoxResult
        '                        msgResult = MsgBox("One or more courses has a course date greater than expiration date. Do you want the system to set the new License Issued date to the first day of the month following the most recent course date", MsgBoxStyle.YesNo, "Licensee Dates")
        '                        If msgResult = MsgBoxResult.Yes Then
        '                            Dim expression1 As String = IIf(mostRecentCourseDate.Month = 12, "01", (mostRecentCourseDate.Month + 1).ToString) + "/" + "01 / " + IIf(mostRecentCourseDate.Month = 12, (mostRecentCourseDate.Year + 1).ToString, mostRecentCourseDate.Year.ToString)
        '                            oLicensee.ISSUED_DATE = CDate(expression1)
        '                        Else
        '                            oLicensee.ISSUED_DATE = oLicensee.LICENSE_EXPIRE_DATE
        '                        End If
        '                    Else
        '                        oLicensee.ISSUED_DATE = oLicensee.LICENSE_EXPIRE_DATE
        '                    End If
        '                    oLicensee.EXCEPT_GRANT_DATE = CDate("01/01/0001")
        '                    '----- resetting the Issued Date complete --- 
        '                    '      now reset the License Expiration Date and print the appropriate forms 
        '                    oLicensee.LICENSE_EXPIRE_DATE = oLicensee.ISSUED_DATE.AddYears(2)
        '                    'Generate Licensee Renewal Letter 
        '                    'Generate the Licensee Certificate corresponding to the certification type
        '                    'Generate the License Card 
        '                    RaiseEvent GenerateLicenseeRenewalLetter()
        '                    RaiseEvent GenerateLicenseeCertLetter()
        '                    RaiseEvent GenerateLicenseeCard()
        '                End If
        '            End If

        '            'If the Status = Renewing, , giving the user the option of generating the Licensee Info Needed letter listing all conditions that prevented the Licensee Status from being set to certified
        '            If oLicensee.StatusDesc = "RENEWING" Then
        '                'If MsgBox("Do you want to generate the licensee info needed letter listing all the conditions that prevented the licensee status from being set to certified").Yes Then
        '                ' generate the licensee info needed letter
        '                'End If
        '            End If

        '            If oLicensee.StatusDesc = "CERTIFIED" Then
        '                oLicensee.EXCEPT_GRANT_DATE = CDate("01/01/0001")
        '            End If

        '            '----- DDD Requirements ------------------------------------------------------------------------
        '            'If the Status prior to the execution of the License Logic was Certified and 
        '            '   is now Not Currently Certified, giving the user the option of generating the Not Certified Letter

        '            '----- changed from CERTIFIED to NOT CURRENTLY CERTIFIED -- 
        '            '      print the Not Certified Letter 
        '            If priorStatus = "CERTIFIED" And oLicensee.StatusDesc = "NOT CURRENTLY CERTIFIED" Then
        '                'If MsgBox("Do you want to generate the NOT CERTIFIED LETTER").Yes Then
        '                'generate the not certified letter
        '                RaiseEvent GenerateNoCertificationLetterOption(oLicensee)
        '                ' End If
        '            End If

        '            '----- changed from APPLICANT to CERTIFIED -------------------
        '            '      or upgrade from CLOSURE to INSTALL     ----------------
        '            '      or changed the hire status (to trigger the printing) --
        '            If (priorStatus = "APPLICANT" And oLicensee.StatusDesc = "CERTIFIED") Or _
        '                (oLicensee.StatusDesc = "CERTIFIED" And priorCertType = "CLOSURE" And oLicenseeInfo.CertTypeDesc = "INSTALL") Or _
        '                (oLicensee.StatusDesc = "CERTIFIED" And (strLicenseeHireStatus <> String.Empty And (strLicenseeHireStatus <> oLicensee.HIRE_STATUS))) Then

        '                'Generate the congratulatory letter
        '                'Generate the License Certificate corresponding to the certification type
        '                'Generate the Licensee Cardv
        '                RaiseEvent GenerateCongLetter()
        '                RaiseEvent GenerateLicenseeCertLetter()
        '                RaiseEvent GenerateLicenseeCard()
        '                oLicensee.ISSUED_DATE = Now.Date
        '                oLicensee.ORIGIN_ISSUED_DATE = Now.Date
        '                Dim expression As String = IIf(Now.Month = 12, "01", (Now.Month + 1).ToString) + "/" + "01/" + (Now.Year + 2).ToString
        '                oLicensee.LICENSE_EXPIRE_DATE = CDate(expression)
        '            End If
        '            oLicensee.OVERRIDE_EXPIRE = False
        '        End If
        '        '-------------------------------------------------------------------------------------------

        '        '----- changed from NOT CURRENTLY CERTIFIED to CERTIFIED -----------------------------------
        '        If (priorStatus = "NOT CURRENTLY CERTIFIED" And oLicensee.StatusDesc = "CERTIFIED") Then
        '            'Generate the congratulatory letter
        '            'Generate the License Certificate corresponding to the certification type
        '            'Generate the Licensee Card
        '            RaiseEvent GenerateCongLetter()
        '            RaiseEvent GenerateLicenseeCertLetter()
        '            RaiseEvent GenerateLicenseeCard()
        '            oLicensee.ISSUED_DATE = Now.Date
        '            Dim expression As String = IIf(Now.Month = 12, "01", (Now.Month + 1).ToString) + "/" + "01/" + (Now.Year + 2).ToString
        '            oLicensee.LICENSE_EXPIRE_DATE = CDate(expression)
        '        End If
        '        '-------------------------------------------------------------------------------------------

        '        '----- changed from NO LONGER WITH COMPANY to CERTIFIED -----------------------------------
        '        If (priorStatus = "NO LONGER WITH COMPANY" And oLicensee.StatusDesc = "CERTIFIED") Then
        '            'Generate the congratulatory letter
        '            'Generate the License Certificate corresponding to the certification type
        '            'Generate the Licensee Card
        '            'Issued date is set to today, and License Expiration Date is not incremented. 
        '            RaiseEvent GenerateCongLetter()
        '            RaiseEvent GenerateLicenseeCertLetter()
        '            RaiseEvent GenerateLicenseeCard()
        '            oLicensee.ISSUED_DATE = Now.Date
        '        End If
        '        '-------------------------------------------------------------------------------------------

        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub

        ''Licensee renewal logic
        'Private Sub renewalLicensee(ByRef oLicensee As MUSTER.Info.LicenseeInfo, ByVal FinRespExpDate As Date, Optional ByVal bolCompany As Boolean = False)
        '    Dim bolFinanciallyResponsible As Boolean
        '    Dim bolInstall As Boolean = False
        '    Dim bolClosure As Boolean = False
        '    Try
        '        If bolCompany Then
        '            If oLicensee.HIRE_STATUS <> "RX - Not for Hire - Owner" Then
        '                If Date.Compare(FinRespExpDate.Date, Now.Date) > 0 Then
        '                    If oLicensee.EMPLOYEE_LETTER Or oLicensee.HIRE_STATUS = "HX - For Hire - Owner" Then
        '                        bolFinanciallyResponsible = True
        '                    Else
        '                        bolFinanciallyResponsible = False
        '                    End If
        '                Else
        '                    bolFinanciallyResponsible = False
        '                End If
        '            Else
        '                bolFinanciallyResponsible = True
        '            End If
        '        Else
        '            bolFinanciallyResponsible = False
        '        End If
        '        Dim xlicCourseInfo As MUSTER.Info.LicenseeCourseInfo
        '        For Each xlicCourseInfo In pLicCourse.colLicCourse.Values
        '            If Date.Compare(xlicCourseInfo.CourseDate, oLicensee.ISSUED_DATE) > 0 And xlicCourseInfo.CourseTypeID = 920 Then
        '                bolClosure = True
        '            End If
        '            If Date.Compare(xlicCourseInfo.CourseDate, oLicensee.ISSUED_DATE) > 0 And xlicCourseInfo.CourseTypeID = 921 Then
        '                bolInstall = True
        '            End If
        '        Next
        '        If bolClosure And bolInstall Then
        '            oLicenseeInfo.CertTypeDesc = "INSTALL"
        '        Else
        '            If oLicenseeInfo.CertTypeDesc = "CLOSURE" Then
        '                If bolClosure Then
        '                    oLicenseeInfo.CertTypeDesc = "CLOSURE"
        '                Else
        '                    oLicensee.StatusDesc = "RENEWING"
        '                End If
        '            Else
        '                oLicensee.StatusDesc = "RENEWING"
        '            End If
        '        End If
        '        If oLicenseeInfo.CertTypeDesc = "INSTALL" Or oLicenseeInfo.CertTypeDesc = "CLOSURE" Then
        '            If bolFinanciallyResponsible Then
        '                If Date.Compare(oLicensee.APP_RECVD_DATE, oLicensee.ISSUED_DATE) > 0 Then
        '                    oLicensee.StatusDesc = "CERTIFIED"
        '                Else
        '                    oLicensee.StatusDesc = "RENEWING"
        '                End If
        '            Else
        '                oLicensee.StatusDesc = "RENEWING"
        '            End If
        '        End If
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        ''Licensee expired logic
        'Private Sub expiredlicensee(ByRef oLicensee As MUSTER.Info.LicenseeInfo, ByVal FinRespExpDate As Date, Optional ByVal bolCompany As Boolean = False)
        '    Dim bolFinanciallyResponsible As Boolean
        '    Dim bolInstall As Boolean = False
        '    Dim bolClosure As Boolean = False
        '    Dim bolInstall1 As Boolean = False
        '    Dim bolClosure1 As Boolean = False
        '    Try
        '        If bolCompany Then
        '            If oLicensee.HIRE_STATUS <> "RX - Not for Hire - Owner" Then
        '                If Date.Compare(FinRespExpDate.Date, Now.Date) > 0 Then
        '                    If oLicensee.EMPLOYEE_LETTER Or oLicensee.HIRE_STATUS = "HX - For Hire - Owner" Then
        '                        bolFinanciallyResponsible = True
        '                    Else
        '                        bolFinanciallyResponsible = False
        '                    End If
        '                Else
        '                    bolFinanciallyResponsible = False
        '                End If
        '            Else
        '                bolFinanciallyResponsible = True
        '            End If
        '        Else
        '            bolFinanciallyResponsible = False
        '        End If
        '        Dim xlicCourseTestInfo As MUSTER.Info.LicenseeCourseTestInfo
        '        For Each xlicCourseTestInfo In pLicCourseTest.colLicCourseTest.Values
        '            If (xlicCourseTestInfo.CourseTypeID = 920 And xlicCourseTestInfo.TestScore > 74 And Date.Compare(xlicCourseTestInfo.TestDate, oLicensee.LICENSE_EXPIRE_DATE) > 0) Then 'xlicCourseTestInfo.TestDate > oLicensee.LICENSE_EXPIRE_DATE) Then

        '                bolInstall = True
        '            End If
        '            If (xlicCourseTestInfo.CourseTypeID = 921 And xlicCourseTestInfo.TestScore > 74 And Date.Compare(xlicCourseTestInfo.TestDate, oLicensee.LICENSE_EXPIRE_DATE) > 0) Then 'xlicCourseTestInfo.TestDate > oLicensee.LICENSE_EXPIRE_DATE) Then
        '                bolClosure = True
        '            End If
        '        Next
        '        Dim xlicCourseInfo As MUSTER.Info.LicenseeCourseInfo
        '        For Each xlicCourseInfo In pLicCourse.colLicCourse.Values
        '            If xlicCourseInfo.CourseTypeID = 920 And Date.Compare(xlicCourseInfo.CourseDate, oLicensee.ISSUED_DATE) > 0 Then 'xlicCourseInfo.CourseDate > oLicensee.LICENSE_EXPIRE_DATE Then
        '                bolInstall1 = True
        '            End If
        '            If xlicCourseInfo.CourseTypeID = 921 And Date.Compare(xlicCourseInfo.CourseDate, oLicensee.ISSUED_DATE) > 0 Then 'xlicCourseInfo.CourseDate > oLicensee.LICENSE_EXPIRE_DATE Then
        '                bolClosure1 = True
        '            End If
        '        Next

        '        '------------------------------------------------------------------------
        '        oLicenseeInfo.CertTypeDesc = "NONE"
        '        If bolClosure And bolClosure1 Then
        '            oLicenseeInfo.CertTypeDesc = "CLOSURE"
        '        End If
        '        If bolClosure And bolInstall Then
        '            If bolClosure1 And bolInstall1 Then
        '                oLicenseeInfo.CertTypeDesc = "INSTALL"
        '            End If
        '        End If
        '        '--------------------------------------------------------------------------

        '        If oLicenseeInfo.CertTypeDesc = "NONE" Then
        '            oLicensee.StatusDesc = "NOT CURRENTLY CERTIFIED"
        '        End If

        '        If oLicenseeInfo.CertTypeDesc = "INSTALL" Or oLicenseeInfo.CertTypeDesc = "CLOSURE" Then
        '            If bolFinanciallyResponsible Then
        '                If Date.Compare(oLicensee.APP_RECVD_DATE, oLicensee.ISSUED_DATE) > 0 Then
        '                    oLicensee.StatusDesc = "CERTIFIED"
        '                Else
        '                    oLicensee.StatusDesc = "NOT CURRENTLY CERTIFIED"
        '                End If
        '            Else
        '                oLicensee.StatusDesc = "NOT CURRENTLY CERTIFIED"
        '            End If
        '        End If

        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Sub
        Public Function GetLicenseesToBeProcessed() As DataSet

            Try
                Return oLicenseeDB.DBGetLicenseesToBeProcessed()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function RetrieveCMList(ByVal facID As Int64) As DataSet
            Try
                Return oLicenseeDB.DBGetCMList(facID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
                Return New DataSet
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Gets all the info
        Function GetAll(ByVal companyID As Integer, Optional ByVal deleted As Boolean = False) As MUSTER.Info.LicenseeCollection
            Try
                colLicensees.Clear()
                ' colManagers.Clear()
                colLicensees = oLicenseeDB.DBGetByCompanyID(companyID)
                ' colManagers = oLicenseeDB.DBManagerGetByCompanyID(companyID)
                Return colLicensees
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetManagerAll(ByVal companyID As Integer, Optional ByVal deleted As Boolean = False) As MUSTER.Info.LicenseeCollection
            Try
                '  colLicensees.Clear()
                colManagers.Clear()
                ' colLicensees = oLicenseeDB.DBGetByCompanyID(companyID)
                colManagers = oLicenseeDB.DBManagerGetByCompanyID(companyID)
                Return colManagers
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oLicenseeInfo = oLicenseeDB.DBGetByID(ID)
                If oLicenseeInfo.ID <= 0 And oLicenseeInfo.ID > -100 Then
                    oLicenseeInfo.ID = nID
                    nID -= 1
                End If
                colLicensees.Add(oLicenseeInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Function Add(ByRef oLicensee As MUSTER.Info.LicenseeInfo) As Boolean
            Try

                If colLicensees.Contains(oLicensee.ID) Then
                    oLicenseeInfo = oLicensee
                Else
                    oLicenseeInfo = oLicensee
                    If oLicenseeInfo.ID <= 0 And oLicenseeInfo.ID > -100 Then
                        oLicenseeInfo.ID = nID
                        nID -= 1
                    End If
                    colLicensees.Add(oLicenseeInfo)
                End If

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oLicenseeInfoLocal As MUSTER.Info.LicenseeInfo
            Try
                For Each oLicenseeInfoLocal In colLicensees.Values
                    If oLicenseeInfoLocal.ID = ID Then
                        colLicensees.Remove(oLicenseeInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Licensee " & ID.ToString & " is not in the collection of Licensees.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLicensee As MUSTER.Info.LicenseeInfo)
            Try
                colLicensees.Remove(oLicensee)
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Licensee " & oLicensee.ID & " is not in the collection of Licensees.")
        End Sub
        Public Sub Remove(ByVal licNumber As Integer, ByVal bol As Boolean)
            Try

                Dim xLicenseeInfo As MUSTER.Info.LicenseeInfo
                For Each xLicenseeInfo In colLicensees.Values
                    If xLicenseeInfo.LICENSEE_NUMBER = licNumber.ToString Then
                        colLicensees.Remove(xLicenseeInfo)
                        Exit Sub
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Try
                Dim xLicenseeInfo As MUSTER.Info.LicenseeInfo
                For Each xLicenseeInfo In colLicensees.Values
                    If xLicenseeInfo.IsDirty Then
                        oLicenseeInfo = xLicenseeInfo
                        Me.Save(moduleID, staffID, returnVal, True)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oLicenseeInfo = New MUSTER.Info.LicenseeInfo
        End Sub
        Public Sub Reset()
            oLicenseeInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oLicenseeInfoLocal As New MUSTER.Info.LicenseeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Dim tempExpDate As String
            Try
                tbEntityTable.Columns.Add("Licensee ID")
                tbEntityTable.Columns.Add("Licensee Name")
                tbEntityTable.Columns.Add("Licensee Number")
                tbEntityTable.Columns.Add("Status")
                tbEntityTable.Columns.Add("Status_ID")
                tbEntityTable.Columns.Add("Certification Type")
                tbEntityTable.Columns.Add("CERT_TYPE_ID")
                tbEntityTable.Columns.Add("Exp. Date")
                tbEntityTable.Columns.Add("Hire Status")

                For Each oLicenseeInfoLocal In colLicensees.Values
                    dr = tbEntityTable.NewRow()
                    dr("Licensee ID") = oLicenseeInfoLocal.ID
                    dr("Licensee Name") = oLicenseeInfoLocal.LAST_NAME + ", " + oLicenseeInfoLocal.FIRST_NAME + IIf(oLicenseeInfoLocal.MIDDLE_NAME = String.Empty, " ", " " + oLicenseeInfoLocal.MIDDLE_NAME)
                    dr("Licensee Number") = oLicenseeInfoLocal.LICENSEE_NUMBER_PREFIX + oLicenseeInfoLocal.LICENSEE_NUMBER.ToString
                    dr("Status") = oLicenseeInfoLocal.StatusDesc
                    dr("Status_ID") = oLicenseeInfoLocal.STATUS
                    dr("Certification Type") = oLicenseeInfoLocal.CertTypeDesc
                    dr("CERT_TYPE_ID") = oLicenseeInfoLocal.CertType
                    'If oLicenseeInfoLocal.LICENSE_EXPIRE_DATE.ToShortDateString = "1/1/0001" Then
                    '    dr("Exp. Date") = String.Empty
                    'Else
                    If oLicenseeInfoLocal.LICENSE_EXPIRE_DATE.Length > 10 Then
                        tempExpDate = oLicenseeInfoLocal.LICENSE_EXPIRE_DATE.Substring(0, 10)
                        dr("Exp. Date") = tempExpDate
                    Else
                        dr("Exp. Date") = oLicenseeInfoLocal.LICENSE_EXPIRE_DATE
                    End If

                    'End If
                    dr("Hire Status") = oLicenseeInfoLocal.HIRE_STATUS
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
 Public Function EntityTableManager() As DataTable
            Dim oManagerInfoLocal As New MUSTER.Info.LicenseeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Dim tempExpDate As String
            Try
                tbEntityTable.Columns.Add("Manager ID")
                tbEntityTable.Columns.Add("Manager Name")
                tbEntityTable.Columns.Add("Manager Number")
                tbEntityTable.Columns.Add("CMStatus")
                tbEntityTable.Columns.Add("CMStatus_ID")
                tbEntityTable.Columns.Add("Certification Type")
                tbEntityTable.Columns.Add("CERT_TYPE_ID")
                tbEntityTable.Columns.Add("Init Cert. By")
                tbEntityTable.Columns.Add("Init Cert. Date")
                tbEntityTable.Columns.Add("Retraining Date1")
                tbEntityTable.Columns.Add("Retraining Date2")
                tbEntityTable.Columns.Add("Retraining Date3")
                tbEntityTable.Columns.Add("Revoke Date")
                tbEntityTable.Columns.Add("RetrainRequiredBy Date")
                tbEntityTable.Columns.Add("Hire Status")

                For Each oManagerInfoLocal In colManagers.Values
                    dr = tbEntityTable.NewRow()
                    dr("Manager ID") = oManagerInfoLocal.ID
                    dr("Manager Name") = oManagerInfoLocal.LAST_NAME + ", " + oManagerInfoLocal.FIRST_NAME + IIf(oManagerInfoLocal.MIDDLE_NAME = String.Empty, " ", " " + oManagerInfoLocal.MIDDLE_NAME)
                    dr("Manager Number") = oManagerInfoLocal.LICENSEE_NUMBER_PREFIX + oManagerInfoLocal.LICENSEE_NUMBER.ToString
                    dr("CMStatus") = oManagerInfoLocal.CMStatusDesc
                    dr("CMStatus_ID") = oManagerInfoLocal.CMSTATUS
                    'dr("Status") = oManagerInfoLocal.StatusDesc
                    ' dr("Status_ID") = oManagerInfoLocal.STATUS
                    dr("Certification Type") = oManagerInfoLocal.CMCertTypeDesc
                    dr("CERT_TYPE_ID") = oManagerInfoLocal.CMCertType
                    dr("Init Cert. By") = oManagerInfoLocal.INITCERTBYDESC
                    'If oLicenseeInfoLocal.LICENSE_EXPIRE_DATE.ToShortDateString = "1/1/0001" Then
                    '    dr("Exp. Date") = String.Empty
                    'Else
                    If oManagerInfoLocal.INITCERTDATE.Length > 10 Then
                        tempExpDate = oManagerInfoLocal.INITCERTDATE.Substring(0, 10)
                        dr("Init Cert. Date") = tempExpDate
                    Else
                        dr("Init Cert. Date") = oManagerInfoLocal.INITCERTDATE
                    End If
                    If oManagerInfoLocal.RetrainDate1.Length > 10 Then
                        tempExpDate = oManagerInfoLocal.RetrainDate1.Substring(0, 10)
                        dr("Retraining Date1") = tempExpDate
                    Else
                        dr("Retraining Date1") = oManagerInfoLocal.RetrainDate1
                    End If
                    If oManagerInfoLocal.RetrainDate2.Length > 10 Then
                        tempExpDate = oManagerInfoLocal.RetrainDate2.Substring(0, 10)
                        dr("Retraining Date2") = tempExpDate
                    Else
                        dr("Retraining Date2") = oManagerInfoLocal.RetrainDate2
                    End If
                    If oManagerInfoLocal.RetrainDate3.Length > 10 Then
                        tempExpDate = oManagerInfoLocal.RetrainDate3.Substring(0, 10)
                        dr("Retraining Date3") = tempExpDate
                    Else
                        dr("Retraining Date3") = oManagerInfoLocal.RetrainDate3
                    End If
                    If oManagerInfoLocal.RevokeDate.Length > 10 Then
                        tempExpDate = oManagerInfoLocal.RevokeDate.Substring(0, 10)
                        dr("Revoke Date") = tempExpDate
                    Else
                        dr("Revoke Date") = oManagerInfoLocal.RevokeDate
                    End If
                    If oManagerInfoLocal.RetrainReqDATE.Length > 10 Then
                        tempExpDate = oManagerInfoLocal.RetrainReqDATE.Substring(0, 10)
                        dr("RetrainRequiredBy Date") = tempExpDate
                    Else
                        dr("RetrainRequiredBy Date") = oManagerInfoLocal.RetrainReqDATE
                    End If
                    'End If
                    dr("Hire Status") = oManagerInfoLocal.HIRE_STATUS
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetLicenseeList(Optional ByVal companyID As Integer = -1) As DataSet
            Dim dsData As DataSet
            Try
                dsData = oLicenseeDB.GetLicenseeList(, companyID)
                Return dsData
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetManagerList(Optional ByVal companyID As Integer = -1) As DataSet
            Dim dsData As DataSet
            Try
                dsData = oLicenseeDB.GetManagerList(, companyID)
                Return dsData
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetPriorCompanies(ByVal LicID As Integer) As DataSet
            Dim dsData As DataSet
            Try
                dsData = oLicenseeDB.GetPriorCompanies(LicID, True)
                Return dsData
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function ProcessRRE() As Boolean
            Try
                Return oLicenseeDB.ProcessRRE()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetLicenseesByType(ByVal _dt As DateTime, Optional ByVal InputType As String = "RENEWAL", Optional ByVal showDeleted As Boolean = False, Optional ByVal showLetterGeneratedOnly As Int16 = -1) As DataSet
            Try
                Return oLicenseeDB.GetLicenseesByType(_dt, InputType, showDeleted, showLetterGeneratedOnly)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub UpdateRenewals(ByVal strLicenseeIDs As String, Optional ByVal InputType As String = "RENEWAL")
            Try
                oLicenseeDB.UpdateRenewals(strLicenseeIDs, InputType)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetLicenseeStatus(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Return oLicenseeDB.DBGetLicenseeStatus(showInActive, showBlankPropertyName)
        End Function
        Public Function GetManagerStatus(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Return oLicenseeDB.DBGetManagerStatus(showInActive, showBlankPropertyName)
        End Function
        Public Function GetLicenseeCertificationType(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Return oLicenseeDB.DBGetLicenseeCertificationType(showInActive, showBlankPropertyName)
        End Function
        Public Function GetManagerInitCertBy(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Return oLicenseeDB.DBGetManagerInitCertBy(showInActive, showBlankPropertyName)
        End Function
        Public Function GetLicenseeHireStatus(Optional ByVal showInActive As Boolean = False, Optional ByVal showBlankPropertyName As Boolean = True) As DataSet
            Return oLicenseeDB.DBGetLicenseeHireStatus(showInActive, showBlankPropertyName)
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub LicenseeInfoChanged(ByVal bolValue As Boolean) Handles oLicenseeInfo.LicenseeInfoChanged
            RaiseEvent LicenseeChanged(bolValue)
        End Sub
        Private Sub LicenseeColChanged(ByVal bolValue As Boolean) Handles colLicensees.LicenseeColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region

        Private Sub pLicCourse_LicenseeCoursesChanged(ByVal bolValue As Boolean) Handles pLicCourse.LicenseeCoursesChanged
            RaiseEvent LicenseeCourseChanged(bolValue)
        End Sub
        Private Sub pLicCourseTest_LicenseeCourseTestsChanged(ByVal bolValue As Boolean) Handles pLicCourseTest.LicenseeCourseTestsChanged
            RaiseEvent LicenseeCourseTestsChanged(bolValue)
        End Sub

    End Class
End Namespace
