'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Company
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      RAF/MKK     05/16/2005  Original class definition
'
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
'                           objects in the internal ReportsCollection.
'
' NOTE: This file to be used as Company to build other objects.
'       Replace keyword "Company" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pCompany
#Region "Public Events"
        Public Event CompanyErr(ByVal MsgStr As String)
        Public Event evtCompanyChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oCompanyInfo As MUSTER.Info.CompanyInfo
        Private WithEvents colCompanys As MUSTER.Info.CompanyCollection
        Private oCompanyDB As New MUSTER.DataAccess.CompanyDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Company").ID
        Private WithEvents oComments As MUSTER.BusinessLogic.pComments
#End Region
#Region "Constructors"
        Public Sub New()
            oCompanyInfo = New MUSTER.Info.CompanyInfo
            colCompanys = New MUSTER.Info.CompanyCollection
            oComments = New MUSTER.BusinessLogic.pComments
        End Sub
        Public Sub New(ByVal CompanyID As Integer)
            oCompanyInfo = New MUSTER.Info.CompanyInfo
            colCompanys = New MUSTER.Info.CompanyCollection
            oComments = New MUSTER.BusinessLogic.pComments
            Me.Retrieve(CompanyID)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return oCompanyInfo.ID
            End Get
            Set(ByVal Value As Int64)
                oCompanyInfo.ID = Value
            End Set
        End Property
        Public Property ACTIVE() As Boolean
            Get
                Return oCompanyInfo.ACTIVE
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.ACTIVE = Value
            End Set
        End Property
        Public Property CE() As Boolean
            Get
                Return oCompanyInfo.CE
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.CE = Value
            End Set
        End Property
        Public Property CM() As Boolean
            Get
                Return oCompanyInfo.CM
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.CM = Value
            End Set
        End Property
        Public Property CERT_RESPON_ID() As String
            Get
                Return oCompanyInfo.CERT_RESPON
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.CERT_RESPON = Value
            End Set
        End Property
        Public Property COMPANY_NAME() As String
            Get
                Return oCompanyInfo.COMPANY_NAME
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.COMPANY_NAME = Value
            End Set
        End Property
        Public Property CREATED_BY() As String
            Get
                Return oCompanyInfo.CREATED_BY
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.CREATED_BY = Value
            End Set
        End Property
        Public Property CTC() As Boolean
            Get
                Return oCompanyInfo.CTC
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.CTC = Value
            End Set
        End Property
        Public Property CTIAC() As Boolean
            Get
                Return oCompanyInfo.CTIAC
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.CTIAC = Value
            End Set
        End Property
        Public Property DATE_CREATED() As Date
            Get
                Return oCompanyInfo.DATE_CREATED
            End Get
            Set(ByVal Value As Date)
                oCompanyInfo.DATE_CREATED = Value
            End Set
        End Property
        Public Property DATE_LAST_EDITED() As Date
            Get
                Return oCompanyInfo.DATE_LAST_EDITED
            End Get
            Set(ByVal Value As Date)
                oCompanyInfo.DATE_LAST_EDITED = Value
            End Set
        End Property
        Public Property DELETED() As Boolean
            Get
                Return oCompanyInfo.DELETED
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.DELETED = Value
            End Set
        End Property
        Public Property EC() As Boolean
            Get
                Return oCompanyInfo.EC
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.EC = Value
            End Set
        End Property
        Public Property ED() As Boolean
            Get
                Return oCompanyInfo.ED
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.ED = Value
            End Set
        End Property
        Public Property EMAIL_ADDRESS() As String
            Get
                Return oCompanyInfo.EMAIL_ADDRESS
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.EMAIL_ADDRESS = Value
            End Set
        End Property
        Public Property ERAC() As Boolean
            Get
                Return oCompanyInfo.ERAC
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.ERAC = Value
            End Set
        End Property
        Public Property FIN_RESP_END_DATE() As Date
            Get
                Return oCompanyInfo.FIN_RESP_END_DATE
            End Get
            Set(ByVal Value As Date)
                oCompanyInfo.FIN_RESP_END_DATE = Value
            End Set
        End Property
        Public Property IRAC() As Boolean
            Get
                Return oCompanyInfo.IRAC
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.IRAC = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oCompanyInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.IsDirty = Value
            End Set
        End Property
        Public Property LAST_EDITED_BY() As String
            Get
                Return oCompanyInfo.LAST_EDITED_BY
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.LAST_EDITED_BY = Value
            End Set
        End Property
        Public Property LDC() As Boolean
            Get
                Return oCompanyInfo.LDC
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.LDC = Value
            End Set
        End Property
        Public Property PRO_ENGIN() As String
            Get
                Return oCompanyInfo.PRO_ENGIN
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.PRO_ENGIN = Value
            End Set
        End Property
        Public Property PRO_ENGIN_ADD_ID() As Integer
            Get
                Return oCompanyInfo.PRO_ENGIN_ADD_ID
            End Get
            Set(ByVal Value As Integer)
                oCompanyInfo.PRO_ENGIN_ADD_ID = Value
            End Set
        End Property
        Public Property PRO_ENGIN_APP_APRV_DATE() As Date
            Get
                Return oCompanyInfo.PRO_ENGIN_APP_APRV_DATE
            End Get
            Set(ByVal Value As Date)
                oCompanyInfo.PRO_ENGIN_APP_APRV_DATE = Value
            End Set
        End Property
        Public Property PRO_ENGIN_LIABIL_DATE() As Date
            Get
                Return oCompanyInfo.PRO_ENGIN_LIABIL_DATE
            End Get
            Set(ByVal Value As Date)
                oCompanyInfo.PRO_ENGIN_LIABIL_DATE = Value
            End Set
        End Property
        Public Property PRO_ENGIN_NUMBER() As String
            Get
                Return oCompanyInfo.PRO_ENGIN_NUMBER
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.PRO_ENGIN_NUMBER = Value
            End Set
        End Property

        Public Property PRO_ENGIN_EMAIL() As String
            Get
                Return oCompanyInfo.PRO_ENGIN_EMAIL
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.PRO_ENGIN_EMAIL = Value
            End Set
        End Property


        Public Property PRO_GEOLO() As String
            Get
                Return oCompanyInfo.PRO_GEOLO
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.PRO_GEOLO = Value
            End Set
        End Property
        Public Property PRO_GEOLO_ADD_ID() As Integer
            Get
                Return oCompanyInfo.PRO_GEOLO_ADD_ID
            End Get
            Set(ByVal Value As Integer)
                oCompanyInfo.PRO_GEOLO_ADD_ID = Value
            End Set
        End Property
        Public Property PRO_GEOLO_NUMBER() As String
            Get
                Return oCompanyInfo.PRO_GEOLO_NUMBER
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.PRO_GEOLO_NUMBER = Value
            End Set
        End Property

        Public Property PRO_GEOLO_EMAIL() As String
            Get
                Return oCompanyInfo.PRO_GEOLO_EMAIL
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.PRO_GEOLO_EMAIL = Value
            End Set
        End Property

        Public Property PTTT() As Boolean
            Get
                Return oCompanyInfo.PTTT
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.PTTT = Value
            End Set
        End Property
        Public Property TL() As Boolean
            Get
                Return oCompanyInfo.TL
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.TL = Value
            End Set
        End Property
        Public Property TST() As Boolean
            Get
                Return oCompanyInfo.TST
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.TST = Value
            End Set
        End Property
        Public Property USTSE() As Boolean
            Get
                Return oCompanyInfo.USTSE
            End Get
            Set(ByVal Value As Boolean)
                oCompanyInfo.USTSE = Value
            End Set
        End Property
        Public Property WL() As Boolean
            Get
                Return oCompanyInfo.WL
            End Get
            Set(ByVal value As Boolean)
                oCompanyInfo.WL = value
            End Set
        End Property
        Public Property colCompany() As MUSTER.Info.CompanyCollection
            Get
                Return colCompanys
            End Get
            Set(ByVal Value As MUSTER.Info.CompanyCollection)
                colCompanys = Value
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
        Public Property CreatedBy() As String
            Get
                Return oCompanyInfo.CREATED_BY
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.CREATED_BY = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oCompanyInfo.DATE_CREATED
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oCompanyInfo.LAST_EDITED_BY
            End Get
            Set(ByVal Value As String)
                oCompanyInfo.LAST_EDITED_BY = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oCompanyInfo.DATE_LAST_EDITED
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.CompanyInfo
            Dim oCompanyInfoLocal As MUSTER.Info.CompanyInfo
            Try
                oCompanyInfo = colCompanys.Item(ID)
                If Not oCompanyInfo Is Nothing Then
                    Return oCompanyInfo
                End If
                oCompanyInfo = oCompanyDB.DBGetByInfoID(ID)
                If oCompanyInfo.ID = 0 Then
                    oCompanyInfo.ID = nID
                    nID -= 1
                End If
                colCompanys.Add(oCompanyInfo)
                Return oCompanyInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal CompanyName As String) As MUSTER.Info.CompanyInfo
            Try
                oCompanyInfo = Nothing
                If colCompanys.Contains(CompanyName) Then
                    oCompanyInfo = colCompanys(CompanyName)
                Else
                    If oCompanyInfo Is Nothing Then
                        oCompanyInfo = New MUSTER.Info.CompanyInfo
                    End If
                    'oCompanyInfo = oCompanyDB.DBGetByName(CompanyName)
                    colCompanys.Add(oCompanyInfo)
                End If
                Return oCompanyInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False) As Boolean
            Dim OldKey As String
            Try
                If Not bolValidated Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not ((oCompanyInfo.ID < 0 And oCompanyInfo.ID > -100) And oCompanyInfo.DELETED) Then
                    OldKey = oCompanyInfo.ID.ToString
                    oCompanyDB.put(oCompanyInfo, moduleID, staffID, returnVal)
                    If oCompanyInfo.ID.ToString <> OldKey Then
                        colCompanys.ChangeKey(OldKey, oCompanyInfo.ID.ToString)
                    End If
                    oCompanyInfo.Archive()
                    oCompanyInfo.IsDirty = False
                End If
                RaiseEvent evtCompanyChanged(oCompanyInfo.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Validates the data before saving
        Public Function ValidateData() As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True
            Try
                If oCompanyInfo.COMPANY_NAME = String.Empty Then
                    errStr += "Please Enter the COMPANY NAME(required)" + vbCrLf
                    validateSuccess = False
                End If
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent CompanyErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        'Gets all the info
        Function GetAll() As MUSTER.Info.CompanyCollection
            Try
                colCompanys.Clear()
                'colCompanys = oCompanyDB.GetAllInfo
                Return colCompanys
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oCompanyInfo = oCompanyDB.DBGetByInfoID(ID)
                If oCompanyInfo.ID = 0 Then
                    oCompanyInfo.ID = nID
                    nID -= 1
                End If
                colCompanys.Add(oCompanyInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Function Add(ByRef oCompany As MUSTER.Info.CompanyInfo) As Boolean
            Try
                oCompanyInfo = oCompany
                If ValidateData() Then
                    If oCompanyInfo.ID = 0 Then
                        oCompanyInfo.ID = nID
                        nID -= 1
                    End If
                    colCompanys.Add(oCompanyInfo)
                    Return True
                Else
                    Return False
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oCompanyInfoLocal As MUSTER.Info.CompanyInfo

            Try
                For Each oCompanyInfoLocal In colCompanys.Values
                    If oCompanyInfoLocal.ID = ID Then
                        colCompanys.Remove(oCompanyInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Company " & ID.ToString & " is not in the collection of Companys.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oCompany As MUSTER.Info.CompanyInfo)
            Try
                colCompanys.Remove(oCompany)
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Company " & oCompany.ID & " is not in the collection of Companys.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xCompanyInfo As MUSTER.Info.CompanyInfo
            For Each xCompanyInfo In colCompanys.Values
                If xCompanyInfo.IsDirty Then
                    oCompanyInfo = xCompanyInfo
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oCompanyInfo = New MUSTER.Info.CompanyInfo
        End Sub
        Public Sub Reset()
            oCompanyInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oCompanyInfoLocal As New MUSTER.Info.CompanyInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("Company ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oCompanyInfoLocal In colCompanys.Values
                    dr = tbEntityTable.NewRow()
                    dr("Company ID") = oCompanyInfoLocal.ID
                    dr("Deleted") = oCompanyInfoLocal.DELETED
                    dr("Created By") = oCompanyInfoLocal.CREATED_BY
                    dr("Date Created") = oCompanyInfoLocal.DATE_CREATED
                    dr("Last Edited By") = oCompanyInfoLocal.LAST_EDITED_BY
                    dr("Date Last Edited") = oCompanyInfoLocal.DATE_LAST_EDITED
                    tbEntityTable.Rows.Add(dr)
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetAssociatedLicensees(ByVal compID As Integer, Optional ByVal deleted As Boolean = False) As DataTable
            Dim dsData As DataSet
            Dim dr As DataRow
            Dim drRow As DataRow
            Dim dtPriorLicensees As New DataTable
            Try
                dsData = oCompanyDB.GetAssociatedLicensees(compID, deleted)
                dtPriorLicensees.Columns.Add("COM_LICENSEE_ID")
                dtPriorLicensees.Columns.Add("Licensee ID")
                dtPriorLicensees.Columns.Add("Licensee Name")
                dtPriorLicensees.Columns.Add("Licensee Number")
                dtPriorLicensees.Columns.Add("Status")
                dtPriorLicensees.Columns.Add("Certification Type")
                dtPriorLicensees.Columns.Add("Exp. Date")
                dtPriorLicensees.Columns.Add("Exp. Grant Date", GetType(Date))
                dtPriorLicensees.Columns.Add("Generate Info Letter")
                dtPriorLicensees.Columns.Add("Generate Not Certified Letter")
                dtPriorLicensees.Columns.Add("Deleted", GetType(Boolean))
                dtPriorLicensees.Columns.Add("Created By")
                dtPriorLicensees.Columns.Add("Date Created")
                dtPriorLicensees.Columns.Add("Edited By")
                dtPriorLicensees.Columns.Add("Date Last Edited")

                For Each drRow In dsData.Tables(0).Rows
                    dr = dtPriorLicensees.NewRow
                    dr("Licensee ID") = drRow("LICENSEE_ID")
                    dr("Licensee Name") = drRow("LAST_NAME") + ", " + drRow("FIRST_NAME") + IIf(drRow("MIDDLE_NAME") = String.Empty, " ", ", " + drRow("MIDDLE_NAME"))
                    dr("Licensee Number") = drRow("LICENSE_NUMBER_PREFIX") + drRow("LICENSE_NUMBER").ToString
                    dr("Status") = drRow("STATUS_DESC")
                    dr("Certification Type") = drRow("CERT_TYPE_DESC")
                    dr("Exp. Date") = drRow("lICENSE_EXPIRE_DATE")
                    dr("Exp. Grant Date") = drRow("EXCEPT_GRANT_DATE")
                    dr("Generate Info Letter") = drRow("GENERATE_INFO_LETTER")
                    dr("Generate Not Certified Letter") = drRow("GENERATE_NOT_CERTIFIED_LETTER")
                    dr("Deleted") = drRow("DELETED")
                    dr("Created By") = drRow("CREATED_BY")
                    dr("Date Created") = drRow("DATE_CREATED")
                    dr("Edited By") = drRow("LAST_EDITED_BY")
                    dr("Date Last Edited") = drRow("DATE_LAST_EDITED")
                    dtPriorLicensees.Rows.Add(dr)
                Next
                Return dtPriorLicensees
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function searchLicensee(Optional ByVal LicenseeName As String = Nothing, Optional ByVal companyName As String = Nothing, Optional ByVal LicenseeAddress As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal erac As Boolean = False, Optional ByVal irac As Boolean = False, Optional ByVal spName As String = Nothing) As DataSet
            Try
                Return oCompanyDB.searchLicensee(LicenseeName, companyName, LicenseeAddress, city, state, erac, irac, spName)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function searchTecCompany(Optional ByVal companyName As String = Nothing, Optional ByVal LicenseeAddress As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal erac As Boolean = False, Optional ByVal irac As Boolean = False) As DataSet
            Try
                Return oCompanyDB.DBsearchTecCompany(companyName, LicenseeAddress, city, state, erac, irac)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        'Private Sub CompanyInfoChanged(ByVal bolValue As Boolean) Handles oCompanyInfo.CompanyInfoChanged
        '    RaiseEvent CompanyChanged(bolValue)
        'End Sub
        Private Sub CompanyColChanged(ByVal bolValue As Boolean) Handles colCompanys.CompanyColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region

        Private Sub oCompanyInfo_CompanyInfoChanged(ByVal bolValue As Boolean) Handles oCompanyInfo.CompanyInfoChanged
            RaiseEvent evtCompanyChanged(bolValue)
        End Sub
    End Class
End Namespace
