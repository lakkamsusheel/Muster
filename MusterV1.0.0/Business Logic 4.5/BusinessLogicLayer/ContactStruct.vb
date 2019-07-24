'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.ContactStruct
'   Provides the operations required to manipulate an ContactStruct object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0     KKM         03/30/05    Class definition.
'   1.1     MR          04/28/05    Added LookUp Functions to Populate Address details
'                                   Added Function to Populate Contact Type.
'                                   Added Active Attribute.

'   1.2    TMF         02/10/09    Testing the check-logic on Contact datatable for empty rows that causes the no object error 
''                                    Modified line 1118
'   1.21   TMF         02/17/09    Found out it also tries to get line 1119 when BOLexist is false. Added another if -  nothing case   


'
' Function          Description
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pContactStruct
#Region "Public Events"
        Public Event evtContactStructErr(ByVal MsgStr As String)
        Public Event CitiesChanged(ByVal dsCities As DataSet)
        Public Event ZipChanged(ByVal dsZip As DataSet)
        Public Event StateChanged(ByVal dsState As DataSet)
        Public Event FipsChanged(ByVal strFIPS As String)
        Public Event CompanyCitiesChanged(ByVal dsCities As DataSet)
        Public Event CompanyZipChanged(ByVal dsZip As DataSet)
        Public Event CompanyStateChanged(ByVal dsState As DataSet)
        Public Event CompanyFipsChanged(ByVal strFIPS As String)
#End Region
#Region "Private Member Variables"
        Private WithEvents oContactStructInfo As MUSTER.Info.ContactStructInfo
        Private WithEvents colContactStruct As MUSTER.Info.ContactStructCollection
        Private WithEvents oContactDatum As MUSTER.BusinessLogic.pContactDatum
        Private oContactStructDB As MUSTER.DataAccess.ContactStructDB
        'Private oEntity As New MUSTER.BusinessLogic.pEntity
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private ContactType As Integer
        Private ContactModule As String
        'Private nEntityTypeID As Integer = oEntity.GetEntity("Contact").ID ' New MUSTER.BusinessLogic.pEntity("Contact").ID
        Private strFIPS As String
        Dim dsGetAll As New DataSet

        '----- copy company address variables -------------
        Private bolCopyCompany As Integer
        Private strCompanyAddrLine1 As String
        Private strCompanyAddrLine2 As String
        Private strCompanyAddrCity As String
        Private strCompanyAddrState As String
        Private strCompanyAddrZipcode As String

#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            oContactStructInfo = New MUSTER.Info.ContactStructInfo
            colContactStruct = New MUSTER.Info.ContactStructCollection
            oContactStructDB = New MUSTER.DataAccess.ContactStructDB
            oContactDatum = New MUSTER.BusinessLogic.pContactDatum(strDBConn, MusterXCEP, oContactStructInfo)
        End Sub
#End Region
#Region "Exposed Attributes"


        Public Property PreferredAddress() As Integer
            Get
                Return oContactStructInfo.PreferredAddress
            End Get
            Set(ByVal Value As Integer)
                oContactStructInfo.PreferredAddress = Value
            End Set
        End Property


        Public Property PreferredAlias() As Integer
            Get
                Return oContactStructInfo.PreferredAlias
            End Get
            Set(ByVal Value As Integer)
                oContactStructInfo.PreferredAlias = Value
            End Set
        End Property


        Public Property childContactID() As Integer
            Get
                Return oContactStructInfo.ChildContactID
            End Get
            Set(ByVal Value As Integer)
                oContactStructInfo.ChildContactID = Value
            End Set
        End Property
        Public ReadOnly Property ContactStructCollection() As MUSTER.Info.ContactStructCollection
            Get
                Return colContactStruct
            End Get
        End Property
        Public ReadOnly Property contactStructInfo() As MUSTER.Info.ContactStructInfo
            Get
                Return oContactStructInfo
            End Get
        End Property
        Public ReadOnly Property ContactAssocID() As Integer
            Get
                Return oContactStructInfo.ContactAssocID
            End Get
        End Property
        Public ReadOnly Property EntityAssocDeleted() As Boolean
            Get
                Return oContactStructInfo.EntityAssocdeleted
            End Get
        End Property
        Public ReadOnly Property EntityAssocActive() As Boolean
            Get
                Return oContactStructInfo.EntityAssocActive
            End Get
        End Property
        Public ReadOnly Property ContactAssocDeleted() As Boolean
            Get
                Return oContactStructInfo.ContactAssocdeleted
            End Get
        End Property
        Public ReadOnly Property ContactAssocActive() As Boolean
            Get
                Return oContactStructInfo.ContactAssocActive
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oContactStructInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oContactStructInfo.IsDirty = Value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim oContactStructInfoLocal As MUSTER.Info.ContactStructInfo
                For Each oContactStructInfoLocal In colContactStruct.Values
                    If oContactStructInfoLocal.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
            End Get
            Set(ByVal Value As Boolean)
            End Set
        End Property
        Public ReadOnly Property ConTypeID() As Integer
            Get
                Return oContactStructInfo.ContactTypeID
            End Get
        End Property
        Public ReadOnly Property ModuleID() As String
            Get
                Return oContactStructInfo.moduleID
            End Get
        End Property
        Public ReadOnly Property ccInfo() As String
            Get
                Return oContactStructInfo.ccInfo
            End Get
        End Property
        Public ReadOnly Property displayAs() As String
            Get
                Return oContactStructInfo.displayAs
            End Get
        End Property
        Public ReadOnly Property entityType() As Integer
            Get
                Return oContactStructInfo.entityType
            End Get
        End Property
        Public ReadOnly Property entityID() As Integer
            Get
                Return oContactStructInfo.EntityID
            End Get
        End Property
        Public ReadOnly Property Active() As Boolean
            Get
                Return oContactStructInfo.EntityAssocActive
            End Get
        End Property
        Public ReadOnly Property ContactDatum() As MUSTER.BusinessLogic.pContactDatum
            Get
                Return oContactDatum
            End Get
        End Property

        Public Property bolCopyCompanyAddress() As Integer
            Get
                Return bolCopyCompany
            End Get
            Set(ByVal Value As Integer)
                bolCopyCompany = Value
            End Set
        End Property

        Public Property strCompAddrLine1() As String
            Get
                Return strCompanyAddrLine1
            End Get
            Set(ByVal Value As String)
                strCompanyAddrLine1 = Value
            End Set
        End Property

        Public Property strCompAddrLine2() As String
            Get
                Return strCompanyAddrLine2
            End Get
            Set(ByVal Value As String)
                strCompanyAddrLine2 = Value
            End Set
        End Property

        Public Property strCompAddrCity() As String
            Get
                Return strCompanyAddrCity
            End Get
            Set(ByVal Value As String)
                strCompanyAddrCity = Value
            End Set
        End Property

        Public Property strCompAddrState() As String
            Get
                Return strCompanyAddrState
            End Get
            Set(ByVal Value As String)
                strCompanyAddrState = Value
            End Set
        End Property

        Public Property strCompAddrZipcode() As String
            Get
                Return strCompanyAddrZipcode
            End Get
            Set(ByVal Value As String)
                strCompanyAddrZipcode = Value
            End Set
        End Property

#End Region

#Region "Exposed Operations"
#Region "Info Operations"
        Public Function GetAll()
            Dim temp As Integer
            Dim ds As New DataSet
            Try

                Dim dsData As New DataSet
                ds = oContactStructDB.DBGetMainDS()
                dsData = oContactStructDB.DBGetContactStruct()
                oContactDatum.GetAll()
                For temp = 0 To (dsData.Tables(0).Rows.Count() - 1)
                    oContactStructInfo = New MUSTER.Info.ContactStructInfo(dsData.Tables(0).Rows(temp))
                    oContactStructInfo.parentContact = oContactDatum.ContactCollection.Item(oContactStructInfo.ParentContactID)
                    If oContactStructInfo.ChildContactID <> 0 Then
                        oContactStructInfo.childContact = oContactDatum.ContactCollection.Item(oContactStructInfo.ChildContactID)
                    End If
                    colContactStruct.Add(oContactStructInfo)
                    oContactStructInfo.ChildContactID = 0
                Next
                If Not dsGetAll Is Nothing Then
                    dsGetAll.Clear()
                End If
                dsGetAll = ds
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function Retrieve(ByVal EntityAssocID As Integer, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ContactStructInfo
            Dim bolDataAged As Boolean = False
            Try
                'If Not (oContactStructInfo.EntityAssocdeleted Or oContactStructInfo.entityAssocID = 0) Then
                'End If
                Dim oContactStructInfoLocal As MUSTER.Info.ContactStructInfo

                If EntityAssocID = 0 Then
                    Add(0)
                Else
                    ' check in collection
                    For Each oContactStructInfoLocal In colContactStruct.Values
                        If oContactStructInfoLocal.entityAssocID = EntityAssocID Then
                            oContactStructInfo = oContactStructInfoLocal
                            Exit Try
                        End If
                    Next
                    ' get by contact id
                    oContactStructInfo = colContactStruct.Item(EntityAssocID.ToString)
                    ' Check for Aged Data here.
                    If Not (oContactStructInfo Is Nothing) Then
                        If oContactStructInfo.IsAgedData = True And oContactStructInfo.IsDirty = False Then
                            bolDataAged = True
                            colContactStruct.Remove(oContactStructInfo)
                        End If
                    End If
                    If oContactStructInfo Is Nothing Or bolDataAged Then
                        Add(EntityAssocID, showDeleted)
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Return oContactStructInfo
        End Function
        Public Function Save(ByVal moduleIdValue As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByVal strMode As String = Nothing, Optional ByVal entityAssID As Integer = 0, Optional ByVal contactAssID As Integer = 0, Optional ByVal EntID As Integer = 0, Optional ByVal entType As Integer = 0, Optional ByVal MODID As Integer = 0, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal bolAssocPersonWithCompany As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                oContactStructInfo.ContactAssocID = contactAssID

                If oContactStructInfo.ContactAssocID <= 0 Then
                    oContactStructInfo.CreatedBy = UserID
                Else
                    oContactStructInfo.modifiedBy = UserID
                End If
                If oContactStructInfo.parentContact.ID <= 0 Then
                    oContactStructInfo.parentContact.CreatedBy = UserID
                Else
                    oContactStructInfo.parentContact.modifiedBy = UserID
                End If

                Dim nContactEntityUpdateFlag As Integer = 0

                If oContactDatum.contactDatumInfo.IsAddressDirty And (strMode = "" Or strMode = "MODIFY") And oContactStructInfo.EntityAssocActive = True Then

                    Dim result As Integer = oContactStructDB.DBModifyContactAddress(oContactStructInfo, oContactStructInfo.parentContact, moduleIdValue, staffID, returnVal, UserID, , nContactEntityUpdateFlag)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    If result = 10 Then  'Associated to Other entities but not with other modules.
                        Dim result1 As MsgBoxResult = MsgBox("Contact is associated with other entity(s) in the module. Do you want to change the address to all the other entity(s) in this module?", MsgBoxStyle.YesNo)

                        If result1 = MsgBoxResult.Yes Then
                            oContactStructDB.DBModifyContactAddress(oContactStructInfo, oContactStructInfo.parentContact, moduleIdValue, staffID, returnVal, UserID, 1, nContactEntityUpdateFlag)
                            If Not returnVal = String.Empty Then
                                Exit Function
                            End If
                            MsgBox("Address change is effected to all the other entity(s) within the module")
                        Else
                            oContactStructDB.DBModifyContactAddress(oContactStructInfo, oContactStructInfo.parentContact, moduleIdValue, staffID, returnVal, UserID, 0, nContactEntityUpdateFlag)
                            If Not returnVal = String.Empty Then
                                Exit Function
                            End If
                            MsgBox("Address change is NOT effected to the other entity(s) in the module")
                        End If
                    ElseIf result = 11 Then ' 'Associated to Other module entities 
                        oContactStructDB.DBModifyContactAddress(oContactStructInfo, oContactStructInfo.parentContact, moduleIdValue, staffID, returnVal, UserID, 0, nContactEntityUpdateFlag)
                        If Not returnVal = String.Empty Then
                            Exit Function
                        End If
                        MsgBox("Contact is associated in Multiple Modules. Please use Reconcilation to Modify the Address for other Module/Entities.")
                    End If
                    'Else
                End If
                If bolAssocPersonWithCompany Then
                    oContactStructInfo.entityAssocID = entityAssID
                End If
                oldID = oContactStructInfo.entityAssocID
                If Not (oContactStructInfo.ContactAssocID > 0) Then
                    oContactStructDB.PutContactRelationship(oContactStructInfo, moduleIdValue, staffID, returnVal, UserID)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If
                    If (strMode = "ADD") Then
                        oContactStructInfo.ChildContactID = 0
                        oContactStructDB.PutContactRelationship(oContactStructInfo, moduleIdValue, staffID, returnVal, UserID)
                        If Not returnVal = String.Empty Then
                            Exit Function
                        End If
                    End If
                End If
                If strMode = "MODIFY" Then
                    oContactStructInfo.entityAssocID = entityAssID
                End If
                If strMode = "ASSOCIATE" Then
                    oContactStructInfo.EntityID = EntID
                    oContactStructInfo.entityType = entType
                    oContactStructInfo.moduleID = MODID
                End If

                If nContactEntityUpdateFlag = 0 Then
                    If Not strMode = "SEARCH" Then
                        If Not strMode = "ADD" Then
                            oContactStructDB.PutEntityContactRelationship(oContactStructInfo, moduleIdValue, staffID, returnVal, UserID)
                            If Not returnVal = String.Empty Then
                                Exit Function
                            End If
                        End If
                    End If
                    If Not bolValidated Then
                        If oldID < 0 Then
                            colContactStruct.ChangeKey(oldID, oContactStructInfo.entityAssocID)
                        End If
                    End If
                End If

                oContactStructInfo.Archive()
                oContactStructInfo.IsDirty = False
                'End If
                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Function ValidateData() As Boolean
            Try
                Dim errStr As String = ""
                Dim validateSuccess As Boolean = True
                If oContactStructInfo.ContactTypeID = 0 Then
                    errStr += "Please enter Contact Type Correctly" + vbCrLf
                    validateSuccess = False
                End If
                If errStr.Length > 0 And Not validateSuccess Then
                    RaiseEvent evtContactStructErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetContactsByEntityAndModule(Optional ByVal nEntityID As Integer = 0, Optional ByVal ModuleID As Integer = 0) As DataSet
            Dim dsContacts As DataSet
            Try
                dsContacts = oContactStructDB.DBGetContactsByEntityAndModule(nEntityID, ModuleID)
                Return dsContacts
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetFilteredContacts(Optional ByVal nEntityID As Integer = 0, Optional ByVal ModuleID As Integer = 0, Optional ByVal strEntityIDs As String = "", Optional ByVal bolActive As Boolean = False, Optional ByVal strentityassocids As String = "", Optional ByVal nEntityType As Integer = 0, Optional ByVal nRelatedEntityType As Integer = 0) As DataSet
            Dim dsContacts As DataSet
            Try
                dsContacts = oContactStructDB.DBGetFilteredContacts(nEntityID, ModuleID, strEntityIDs, bolActive, strentityassocids, nEntityType, nRelatedEntityType)
                Return dsContacts
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetContactEntityAssociation(Optional ByVal nEntityAssocID As Integer = 0, Optional ByVal nModuleID As Integer = 0) As MUSTER.Info.ContactStructInfo
            Dim dsConEntityAssoc As DataSet
            Dim temp As Integer
            Try
                dsConEntityAssoc = oContactStructDB.DBGetContactStruct(nEntityAssocID, nModuleID)
                For temp = 0 To (dsConEntityAssoc.Tables(0).Rows.Count() - 1)
                    oContactStructInfo = New MUSTER.Info.ContactStructInfo(dsConEntityAssoc.Tables(0).Rows(temp))
                    'oContactStructInfo.parentContact = oContactDatum.ContactCollection.Item(oContactStructInfo.ParentContactID)
                    'If oContactStructInfo.ChildContactID <> 0 Then
                    '    oContactStructInfo.childContact = oContactDatum.ContactCollection.Item(oContactStructInfo.ChildContactID)
                    'End If
                    'colContactStruct.Add(oContactStructInfo)
                    'oContactStructInfo.ChildContactID = 0
                Next
                Return oContactStructInfo
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Public Sub Add(ByVal id As Integer, Optional ByVal showDeleted As Boolean = False)
            Try
                oContactStructInfo = oContactStructDB.DBGetInfoByID(id, showDeleted)
                If oContactStructInfo.ContactAssocID = 0 Then
                    oContactStructInfo.ContactAssocID = nID
                    nID -= 1
                End If
                colContactStruct.Add(oContactStructInfo)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Add(ByRef ContactStructInfo As MUSTER.Info.ContactStructInfo)
            Try
                oContactStructInfo = ContactStructInfo
                colContactStruct.Add(oContactStructInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Remove(ByVal id As Integer, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim success As Boolean
            Try
                oContactStructInfo = Me.GetContactEntityAssociation(id)
                oContactStructInfo.parentContact = ContactDatum.Retrieve(oContactStructInfo.ParentContactID)
                If oContactStructInfo.ChildContactID <> 0 Then
                    oContactStructInfo.childContact = ContactDatum.Retrieve(oContactStructInfo.ChildContactID)
                End If
                oContactStructInfo.EntityAssocdeleted = True
                success = Me.Save(moduleID, staffID, returnVal, UserID, Nothing, oContactStructInfo.entityAssocID, oContactStructInfo.ContactAssocID, oContactStructInfo.EntityID, oContactStructInfo.entityType, oContactStructInfo.moduleID, True, True)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
                If success Then
                    'MsgBox("successfully deleted")
                    Exit Sub
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Throw New Exception("Entity Assoc ID " & id.ToString & " is not in the collection of Contact Struct.")
        End Sub
        Public Sub Remove(ByVal oContactStructInf As MUSTER.Info.ContactStructInfo)
            Try
                colContactStruct.Remove(oContactStructInf)
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Contact Assoc ID " & oContactStructInf.ContactAssocID & " is not in the collection of Contact Struct.")
        End Sub
        Public Sub Flush()
            Try
            Catch ex As Exception
            End Try
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oContactStructInfo = New MUSTER.Info.ContactStructInfo
        End Sub
        Public Sub Reset()
            oContactStructInfo.Reset()
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colContactStruct.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ContactAssocID.ToString))
            If colIndex + direction > -1 Then
                If colIndex + direction <= nArr.GetUpperBound(0) Then
                    Return colContactStruct.Item(nArr.GetValue(colIndex + direction)).ContactAssocID.ToString
                Else
                    Return colContactStruct.Item(nArr.GetValue(0)).ContactAssocID.ToString
                End If
            Else
                Return colContactStruct.Item(nArr.GetValue(nArr.GetUpperBound(0))).ContactAssocID.ToString
            End If
        End Function
#End Region
#Region "LookUp Operations"
        Public Function PopulateEntityCode() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vORG_ENTITY_TYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateContactType() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCONTACT_TYPE")
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub putContactType(ByVal contactType As String, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByVal contactTypeID As Integer = 0, Optional ByVal deleted As Integer = 0, Optional ByVal letterContactType As Integer = 0)
            Try
                oContactStructDB.PutContactTypesAdmin(contactTypeID, contactType, moduleID, deleted, letterContactType, staffID, returnVal, UserID)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetModuleID(ByVal strModuleName As String) As Integer
            Try
                Dim dtReturn As DataTable = GetDataTable("vMODULENAME", strModuleName)
                Dim drow As DataRow
                If dtReturn.Rows.Count > 0 Then
                    For Each drow In dtReturn.Rows
                        Return drow("Property_ID")
                        Exit Function
                    Next
                End If
                Return 0
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Private Function GetDataTable(ByVal DBViewName As String, Optional ByVal strModuleName As String = "") As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                If strModuleName = String.Empty Then
                    strSQL = "SELECT * FROM " & DBViewName
                Else
                    strSQL = "SELECT * FROM " & DBViewName & " WHERE PROPERTY_NAME = '" & strModuleName & "'"
                End If

                dsReturn = oContactStructDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                Else
                    dtReturn = Nothing
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateAddressDetails(Optional ByVal strCity As String = "", Optional ByVal strState As String = "", Optional ByVal strZip As String = "")

            Dim strWhere As String = String.Empty
            Dim strMessage As String = ""
            Dim dsTemp As DataSet
            Try

                If strCity <> String.Empty Then
                    strWhere += " AND UPPER(CITY) = '" & strCity.ToUpper & "'"
                End If

                If strState <> String.Empty Then
                    strWhere += " AND UPPER(STATE) = '" & strState.ToUpper & "'"
                End If

                If strZip <> String.Empty Then
                    strWhere += " AND ZIP = '" & strZip & "'"
                End If

                If strWhere <> String.Empty Then
                    strWhere = " WHERE " & strWhere.Substring(5, strWhere.Length - 5) & " "
                End If

                If strState = String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT STATE from tblSYS_ZIPCODES " & strWhere & " ORDER BY STATE")
                    RaiseEvent StateChanged(dsTemp)
                    dsTemp.Tables.Clear()
                End If

                If strState = String.Empty And strCity = String.Empty And strZip = String.Empty Then
                    strWhere = " WHERE UPPER(STATE) = 'MS'"
                End If

                If strCity = String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT CITY from tblSYS_ZIPCODES " & strWhere & " ORDER BY CITY")
                    RaiseEvent CitiesChanged(dsTemp)
                    dsTemp.Tables.Clear()
                End If

                If strZip = String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT ZIP  from tblSYS_ZIPCODES " & strWhere & "  ORDER BY ZIP")
                    RaiseEvent ZipChanged(dsTemp)
                    dsTemp.Tables.Clear()
                End If

                If strState <> String.Empty And strCity <> String.Empty And strZip <> String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT FIPS from tblSYS_ZIPCODES " & strWhere & " ORDER BY FIPS")
                    strFIPS = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("FIPS")), "", dsTemp.Tables(0).Rows(0).Item("FIPS"))
                    RaiseEvent FipsChanged(strFIPS)
                    dsTemp.Tables.Clear()
                End If

                If strMessage.Length > 0 Then
                    Throw New Exception("Please Enter Valid Data.")
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateAddressLookUp(ByVal strSQL As String) As DataSet
            Dim dsReturn As New DataSet

            Try
                dsReturn = oContactStructDB.DBGetDS(strSQL)
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function CompanyPopulateAddressDetails(Optional ByVal strCity As String = "", Optional ByVal strState As String = "", Optional ByVal strZip As String = "")
            Dim strWhere As String = String.Empty
            Dim strMessage As String = ""
            Dim dsTemp As DataSet
            Try

                If strCity <> String.Empty Then
                    strWhere += " AND UPPER(CITY) = '" & strCity.ToUpper & "'"
                End If

                If strState <> String.Empty Then
                    strWhere += " AND UPPER(STATE) = '" & strState.ToUpper & "'"
                End If

                If strZip <> String.Empty Then
                    strWhere += " AND ZIP = '" & strZip & "'"
                End If

                If strWhere <> String.Empty Then
                    strWhere = " WHERE " & strWhere.Substring(5, strWhere.Length - 5) & " "
                End If

                If strState = String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT STATE from tblSYS_ZIPCODES " & strWhere & " ORDER BY STATE")
                    RaiseEvent CompanyStateChanged(dsTemp)
                    dsTemp.Tables.Clear()
                End If

                If strState = String.Empty And strCity = String.Empty And strZip = String.Empty Then
                    strWhere = " WHERE UPPER(STATE) = 'MS'"
                End If

                If strCity = String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT CITY from tblSYS_ZIPCODES " & strWhere & " ORDER BY CITY")
                    RaiseEvent CompanyCitiesChanged(dsTemp)
                    dsTemp.Tables.Clear()
                End If

                If strZip = String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT ZIP  from tblSYS_ZIPCODES " & strWhere & "  ORDER BY ZIP")
                    RaiseEvent CompanyZipChanged(dsTemp)
                    dsTemp.Tables.Clear()
                End If

                If strState <> String.Empty And strCity <> String.Empty And strZip <> String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT FIPS from tblSYS_ZIPCODES " & strWhere & " ORDER BY FIPS")
                    strFIPS = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("FIPS")), "", dsTemp.Tables(0).Rows(0).Item("FIPS"))
                    RaiseEvent CompanyFipsChanged(strFIPS)
                    dsTemp.Tables.Clear()
                End If

                If strMessage.Length > 0 Then
                    Throw New Exception("Please Enter Valid Data.")
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oContactStructInfoLocal As New MUSTER.Info.ContactStructInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("ContactAssocID")
                tbEntityTable.Columns.Add("Parent_Contact")
                tbEntityTable.Columns.Add("Child_Contact")

                For Each oContactStructInfoLocal In colContactStruct.Values
                    dr = tbEntityTable.NewRow()
                    dr("ContactAssocID") = oContactStructInfoLocal.ContactAssocID
                    dr("Parent_Contact") = oContactStructInfoLocal.parentContact
                    dr("Child_Contact") = oContactStructInfoLocal.childContact
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function SearchContact(Optional ByVal contactName As String = Nothing, Optional ByVal address As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal phone1 As String = Nothing, Optional ByVal phone2 As String = Nothing, Optional ByVal cell As String = Nothing, Optional ByVal fax As String = Nothing, Optional ByVal email As String = Nothing, Optional ByVal spName As String = "") As DataSet
            Try
                Dim ds As DataSet
                ds = oContactStructDB.DBGetSearchContact(contactName, address, city, state, phone1, phone2, cell, fax, email, spName)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetChildContacts(ByVal parentContactID As Integer) As DataSet
            Try
                Dim ds As DataSet
                ds = oContactStructDB.DBGetChildContacts(parentContactID)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function GetContactAliases(ByVal ContactID As Integer, Optional ByVal comboBox As Boolean = False) As DataSet
            Try
                Dim ds As DataSet
                ds = oContactStructDB.DBGetContactAliases(ContactID, comboBox)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function GetContactAliasAddresses(ByVal ContactID As Integer, Optional ByVal comboBox As Boolean = False) As DataSet
            Try
                Dim ds As DataSet
                ds = oContactStructDB.DBGetContactAddresses(ContactID, comboBox)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function RemoveContactAlias(ByVal ContactID As Integer, ByVal contactAliasID As Integer) As DataSet
            Try

                oContactStructDB.DBPutContactAlias(0, String.Empty, 1, String.Empty, contactAliasID)
                Return GetContactAliases(ContactID)

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function AddContactAlias(ByVal aliasName As String, ByVal contactID As Integer, ByVal userID As String) As DataSet

            Try
                oContactStructDB.DBPutContactAlias(contactID, aliasName, 0, userID)
                Return GetContactAliases(contactID)

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function RemoveContactAddresses(ByVal ContactID As Integer, ByVal addressID As Integer) As DataSet
            Try

                oContactStructDB.DBRemoveContactAddresses(ContactID, addressID)
                Return GetContactAliasAddresses(ContactID)

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function SetContactAddressesToMain(ByVal ContactID As Integer, ByVal addressID As Integer) As DataSet
            Try

                oContactStructDB.DBSetContactAddressesAsMainAddress(ContactID, addressID)
                Return GetContactAliasAddresses(ContactID)

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function

        Public Function getReconciliation(Optional ByVal nOldContactAssocID As Integer = 0) As DataSet
            Dim ds As DataSet
            Try
                ds = oContactStructDB.getReconciliation(nOldContactAssocID)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub UpdateReconciliation(ByVal strAccept As String, ByVal strReject As String, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Try
                oContactStructDB.UpdateReconciliation(strAccept, strReject, moduleID, staffID, returnVal, UserID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function getModules(Optional ByVal contactName As String = Nothing, Optional ByVal address As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal phone1 As String = Nothing, Optional ByVal phone2 As String = Nothing, Optional ByVal cell As String = Nothing, Optional ByVal fax As String = Nothing, Optional ByVal email As String = Nothing, Optional ByVal spName As String = "") As DataSet
            Dim strSQL As String
            Dim ds As DataSet
            Try
                strSQL = "select property_id as ModuleID,property_name as ModuleName from tblsys_property_master where property_type_id = 89"
                ds = oContactStructDB.DBGetDS(strSQL)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function getLetterContactType(Optional ByVal contactName As String = Nothing, Optional ByVal address As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal phone1 As String = Nothing, Optional ByVal phone2 As String = Nothing, Optional ByVal cell As String = Nothing, Optional ByVal fax As String = Nothing, Optional ByVal email As String = Nothing, Optional ByVal spName As String = "") As DataSet
            Dim strSQL As String
            Dim ds As DataSet
            Try
                strSQL = "select property_id as LetterContactTypeID,property_name as LetterContactTypeName from tblsys_property_master where property_type_id = 147"
                ds = oContactStructDB.DBGetDS(strSQL)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function getContactTypes(Optional ByVal contactName As String = Nothing, Optional ByVal address As String = Nothing, Optional ByVal city As String = Nothing, Optional ByVal state As String = Nothing, Optional ByVal phone1 As String = Nothing, Optional ByVal phone2 As String = Nothing, Optional ByVal cell As String = Nothing, Optional ByVal fax As String = Nothing, Optional ByVal email As String = Nothing, Optional ByVal spName As String = "") As DataSet
            Dim strSQL As String
            Dim ds As DataSet
            Try
                strSQL = "select * from vCON_ContactTypes"
                ds = oContactStructDB.DBGetDS(strSQL)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function getContactAddress(ByVal nContactID As Integer) As DataTable
            Dim strSQL As String
            Dim ds As DataSet
            Try
                ds = oContactStructDB.DBGetDS("spCONContactGetCompanyAddress", nContactID)
                Return ds.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function getStates() As DataSet
            Dim strSQL As String
            Dim ds As DataSet
            Try
                strSQL = "SELECT DISTINCT STATE FROM tblSYS_ZIPCODES where state is not null"
                ds = oContactStructDB.DBGetDS(strSQL)
                Return ds
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function FilterContactTypes(ByVal nEntityID As Integer, ByVal nEntityType As Integer, ByVal strModule As String, ByVal nModuleID As Integer, Optional ByVal nContactID As Integer = 0, Optional ByVal bolActive As Boolean = True) As DataTable

            Dim dtContactTypes As DataTable
            Dim dsContacts As DataSet
            Dim nContactXL As Integer = 0
            Dim nContactXH As Integer = 0
            Dim strFilter As String = String.Empty
            Dim drContacts() As DataRow
            Dim i As Integer = 0
            Dim CurrentContactXH As Boolean = False
            Dim CurrentContactXL As Boolean = False
            Try

                dsContacts = oContactStructDB.DBGetMainDS()
                strFilter = "MODULEID = " + nModuleID.ToString + " AND ENTITYID=" + nEntityID.ToString
                drContacts = dsContacts.Tables(0).Select(strFilter)
                If nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.ClosureEvent Or nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.LUST_Event Or nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.FinancialEvent Then
                    strFilter = "ENTITYTYPE = " + nEntityType.ToString + " AND ENTITYID=" + nEntityID.ToString + " AND ACTIVE=" + IIf(bolActive, "1", "0")
                    drContacts = dsContacts.Tables(0).Select(strFilter)
                End If
                dtContactTypes = PopulateContactType()
                For i = 0 To UBound(drContacts)
                    If drContacts(i).Item("LetterContactType") = 1185 Then 'And nContactID <> drContacts(i).Item("ContactID") Then
                        nContactXH = Integer.Parse(drContacts(i).Item("LetterContactType"))
                        If nContactID = drContacts(i).Item("ContactID") Then
                            CurrentContactXH = True
                        End If
                    ElseIf drContacts(i).Item("LetterContactType") = 1186 Then 'And nContactID <> drContacts(i).Item("ContactID") Then
                        nContactXL = Integer.Parse(drContacts(i).Item("LetterContactType"))
                        If nContactID = drContacts(i).Item("ContactID") Then
                            CurrentContactXL = True
                        End If
                    End If
                Next
                Return FilterChildern(dtContactTypes, strModule, nModuleID, nEntityType, nContactXH, nContactXL, CurrentContactXH, CurrentContactXL)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Function FilterChildern(ByVal dtContactTypes As DataTable, ByVal strModule As String, ByVal nModuleID As Integer, ByVal nEntityType As Integer, Optional ByVal nLetterTypeXH As Integer = 0, Optional ByVal nLetterTypeXL As Integer = 0, Optional ByVal CurrentContactXH As Boolean = False, Optional ByVal CurrentContactXL As Boolean = False) As DataTable
            Dim dtTable As New DataTable
            Dim drRow As DataRow
            Dim bolExists As Boolean
            Dim strFilter As String = String.Empty
            Dim i As Integer = 0
            Dim drContactTypes() As DataRow
            'Dim strEntity As String = String.Empty
            Dim nEntityTypeID As Integer
            Dim bolOwnerXL As Boolean = False
            Try
                strFilter = "MODULEID = " + nModuleID.ToString
                drContactTypes = dtContactTypes.Select(strFilter)
                dtTable.Columns.Add("CONTACTTYPE")
                dtTable.Columns.Add("CONTACTTYPEID")
                dtTable.Columns.Add("moduleid")
                dtTable.Columns.Add("LETTERCONTACTTYPE")
                For i = 0 To UBound(drContactTypes)
                    bolExists = False
                    If strModule.ToUpper = "registration".ToUpper Then
                        If nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Owner Or nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Facility Then
                            If nLetterTypeXL = EnumContactType.XL Then
                                If Integer.Parse(drContactTypes(i).Item("LETTERCONTACTTYPE")) <> nLetterTypeXL Then
                                    bolExists = True
                                End If
                                If (CurrentContactXL And Integer.Parse(drContactTypes(i).Item("LETTERCONTACTTYPE")) = nLetterTypeXL) Then
                                    bolExists = True
                                End If
                            End If
                        ElseIf nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Facility Then
                            If Integer.Parse(drContactTypes(i).Item("LETTERCONTACTTYPE")) = EnumContactType.X Then
                                bolExists = True
                            End If
                        End If
                    ElseIf strModule.ToUpper = "CLOSURE".ToUpper Or strModule.ToUpper = "TECHNICAL" Or strModule.ToUpper = "Financial".ToUpper Or strModule.ToUpper = "fees".ToUpper Or strModule.ToUpper = "company".ToUpper Or strModule.ToUpper = "C & E".ToUpper Then

                        Select Case strModule.ToUpper
                            Case "closure".ToUpper
                                If nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Owner Then
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.Owner
                                ElseIf nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Facility Then
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.Facility
                                Else
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.ClosureEvent
                                End If
                            Case "technical".ToUpper
                                If nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Owner Then
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.Owner
                                ElseIf nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Facility Then
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.Facility
                                Else
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.LUST_Event
                                End If
                            Case "financial".ToUpper
                                If nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Owner Then
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.Owner
                                ElseIf nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Facility Then
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.Facility
                                Else
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.FinancialEvent
                                End If
                            Case "fees".ToUpper, "company".ToUpper, "C & E".ToUpper
                                If nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Owner Then
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.Owner
                                Else
                                    nEntityTypeID = oContactStructDB.SqlHelperProperty.EntityTypes.Facility
                                End If
                        End Select

                        bolExists = FilterContactTypeForEvents(drContactTypes(i), nEntityType, bolExists, nLetterTypeXH, nLetterTypeXL, nEntityTypeID, CurrentContactXH, CurrentContactXL)
                    End If

                    If bolExists = True Then
                        drRow = dtTable.NewRow
                        drRow("CONTACTTYPE") = drContactTypes(i).Item("CONTACTTYPE")
                        drRow("CONTACTTYPEID") = drContactTypes(i).Item("CONTACTTYPEID")
                        drRow("moduleid") = drContactTypes(i).Item("moduleid")
                        drRow("LETTERCONTACTTYPE") = drContactTypes(i).Item("LETTERCONTACTTYPE")
                        dtTable.Rows.Add(drRow)
                    End If
                Next
                If dtTable.Rows.Count > 0 Then
                    Return dtTable
                Else
                    Return dtContactTypes
                End If

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Enum EnumContactType
            XH = 1185
            XL = 1186
            X = 1184
        End Enum
        Private Function FilterContactTypeForEvents(ByVal drContactTypes As DataRow, ByVal nEntityType As Integer, ByVal bolExists As Boolean, ByVal nLetterTypeXH As Integer, ByVal nLettertypeXL As Integer, ByVal nEntityTypeID As Integer, Optional ByVal CurrentContactXH As Boolean = False, Optional ByVal CurrentContactXL As Boolean = False) As Boolean
            If nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Owner Or _
                        nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.Facility Then
                If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) = EnumContactType.X Then
                    bolExists = True
                End If
                If nLettertypeXL = EnumContactType.XL And nLetterTypeXH = 0 Then
                    If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) <> nLettertypeXL Then
                        bolExists = True
                    End If
                End If
                If nLetterTypeXH = EnumContactType.XH And nLettertypeXL = 0 Then
                    If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) <> nLetterTypeXH Then
                        bolExists = True
                    End If
                End If
                'ElseIf nEntityType = oEntity.GetEntity("Facility").ID Then
                '    If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) = EnumContactType.X Then
                '        bolExists = True
                '    End If
                '    If nLettertypeXL = EnumContactType.XL And nLetterTypeXH = 0 Then
                '        If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) <> nLettertypeXL Then
                '            bolExists = True
                '        End If
                '    End If
                '    If nLetterTypeXH = EnumContactType.XH And nLettertypeXL = 0 Then
                '        If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) <> nLetterTypeXH Then
                '            bolExists = True
                '        End If
                '    End If
            ElseIf nEntityType = nEntityTypeID Then
                If nLettertypeXL = EnumContactType.XL And nLetterTypeXH = 0 Then
                    If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) <> nLettertypeXL Then
                        bolExists = True
                    End If
                End If
                If nLetterTypeXH = EnumContactType.XH And nLettertypeXL = 0 Then
                    If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) <> nLetterTypeXH Then
                        bolExists = True
                    End If
                End If
            End If

            If Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) = EnumContactType.X Or (nLettertypeXL = 0 And nLetterTypeXH = 0) Then
                bolExists = True
            End If

            If (CurrentContactXH And Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) = nLetterTypeXH) Then
                bolExists = True
            End If

            If (CurrentContactXL And Integer.Parse(drContactTypes.Item("LETTERCONTACTTYPE")) = nLettertypeXL) Then
                bolExists = True
            End If


            ',CurrentContactXL

            Return bolExists
        End Function

        Public Function GETContactName(ByVal nEntityID As Integer, ByVal nEntityType As Integer, ByVal nModuleID As Integer, Optional ByVal nOwnerID As Integer = 0, Optional ByVal nActive As Integer = -1) As DataTable
            Dim dsContacts As DataSet
            Dim drContacts() As DataRow
            Dim i As Integer = 0
            Dim strFilter As String = String.Empty
            Dim dtTable As New DataTable
            Dim drRow As DataRow
            Dim bolExist As Boolean = False
            Try
                dtTable.Columns.Add("CONTACT_Name")
                dtTable.Columns.Add("Greeting")
                dtTable.Columns.Add("EntityID")
                dtTable.Columns.Add("Type")
                dtTable.Columns.Add("Address_One")
                dtTable.Columns.Add("Address_Two")
                dtTable.Columns.Add("State")
                dtTable.Columns.Add("City")
                dtTable.Columns.Add("Zip")
                dtTable.Columns.Add("Phone")
                dtTable.Columns.Add("ContactAssocID")
                dtTable.Columns.Add("AssocCompany")
                dtTable.Columns.Add("IsPerson")
                dtTable.Columns.Add("ContactType")
                dtTable.Columns.Add("ContactID")
                dtTable.Columns.Add("Vendor_Number")
                dtTable.Columns.Add("First_Name")
                dtTable.Columns.Add("Last_Name")
                dtTable.Columns.Add("TITLE")

                dsContacts = oContactStructDB.DBGetMainDS()
                If nModuleID = 612 Then
                    strFilter = "MODULEID = " + nModuleID.ToString + " AND ENTITYID=" + nEntityID.ToString
                End If
                If nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.ClosureEvent Or nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.LUST_Event Or nEntityType = oContactStructDB.SqlHelperProperty.EntityTypes.FinancialEvent Then
                    strFilter += "((MODULEID = " + nModuleID.ToString + " AND ENTITYTYPE = " + nEntityType.ToString + " AND ENTITYID=" + nEntityID.ToString + ") OR ( MODULEID =612  AND ENTITYID=" + nOwnerID.ToString + " AND LetterContactType = 1186))"
                End If
                If nActive = 0 Or nActive = 1 Then
                    strFilter += " AND ACTIVE = " + nActive.ToString
                End If
                If dsContacts.Tables.Count > 0 Then
                    drContacts = dsContacts.Tables(0).Select(strFilter)

                    'Check to see if Contacts return a datatable as well as have any existing rows
                    If Not drContacts Is Nothing AndAlso drContacts.GetUpperBound(0) >= 0 Then
                        Dim vendorNumStr As String
                        vendorNumStr = ""
                        For i = 0 To UBound(drContacts)
                            bolExist = False
                            If drContacts(i).Item("LetterContactType") = EnumContactType.XH Then
                                bolExist = True
                            ElseIf drContacts(i).Item("LetterContactType") = EnumContactType.XL Then
                                bolExist = True
                            End If
                            If bolExist Or UCase(drContacts(i).Item("TYPE")) = "FINANCIAL PAYEE" Then
                                drRow = dtTable.NewRow
                                drRow("CONTACT_Name") = drContacts(i).Item("CONTACT_Name")
                                drRow("EntityID") = drContacts(i).Item("ENTITYID")
                                drRow("Type") = drContacts(i).Item("LetterContactType")
                                drRow("Address_One") = drContacts(i).Item("Address_One")
                                If Not drContacts(i).Item("Address_Two") Is System.DBNull.Value Then
                                    drRow("Address_Two") = drContacts(i).Item("Address_Two")
                                Else
                                    drRow("Address_Two") = ""
                                End If
                                drRow("City") = drContacts(i).Item("City")
                                drRow("State") = drContacts(i).Item("State")
                                drRow("Zip") = drContacts(i).Item("Zip")
                                If Not drContacts(i).Item("Phone_Number_One") Is System.DBNull.Value Then
                                    drRow("Phone") = drContacts(i).Item("Phone_Number_One")
                                Else
                                    drRow("Phone") = ""
                                End If
                                drRow("ContactAssocID") = drContacts(i).Item("ContactAssocID")
                                drRow("AssocCompany") = drContacts(i).Item("AssocCompany")
                                drRow("IsPerson") = drContacts(i).Item("IsPerson")
                                drRow("ContactType") = drContacts(i).Item("TYPE")
                                drRow("ContactID") = drContacts(i).Item("ContactID")
                                If Not drContacts(i).Item("Vendor_Number") Is System.DBNull.Value Then
                                    drRow("Vendor_Number") = drContacts(i).Item("Vendor_Number")
                                    If vendorNumStr.Length <= 0 Then
                                        vendorNumStr = drRow("Vendor_Number")
                                    End If
                                Else
                                    drRow("Vendor_Number") = ""
                                End If
                                If Not drContacts(i).Item("Last_Name") Is System.DBNull.Value Then
                                    drRow("First_Name") = drContacts(i).Item("First_Name")
                                    drRow("Last_Name") = drContacts(i).Item("Last_Name")
                                    drRow("TITLE") = drContacts(i).Item("TITLE")

                                    If drRow("TITLE") <> String.Empty Then
                                        drRow("Greeting") = String.Format("{0} {1}", drContacts(i).Item("TITLE"), drContacts(i).Item("Last_Name"))
                                    Else
                                        drRow("Greeting") = String.Empty

                                    End If

                                Else
                                    drRow("First_Name") = ""
                                    drRow("Last_Name") = ""
                                    drRow("TITLE") = ""
                                End If
                                dtTable.Rows.Add(drRow)
                            End If
                        Next

                        If Not drRow Is Nothing Then
                            drRow("Vendor_Number") = vendorNumStr
                        End If

                    End If
                End If
                Return dtTable
            Catch ex As Exception
                Throw ex
            End Try

        End Function
        Public Function GetContactsForAllModules(ByVal nEntityID As Integer, Optional ByVal nEntities As String = "") As String
            Dim strAllContactsIdTags As String
            Dim drRow As DataRow
            Dim Str As String = String.Empty
            Dim dsSet As DataSet
            Dim rowcount As Integer = 0

            Try
                strAllContactsIdTags = String.Empty
                dsSet = oContactStructDB.DBGetContactsForAllModules(nEntityID, nEntities)
                If dsSet.Tables.Count > 0 Then
                    If dsSet.Tables(0).Rows.Count > 0 Then
                        For Each drRow In dsSet.Tables(0).Rows
                            If rowcount < dsSet.Tables(0).Rows.Count - 1 Then
                                Str = ","
                            Else
                                Str = ""
                            End If
                            If Not drRow("CONTACTASSOCID") Is System.DBNull.Value Then
                                strAllContactsIdTags += drRow("CONTACTASSOCID").ToString + Str
                            End If
                            rowcount += 1
                        Next
                    End If
                End If
                Return strAllContactsIdTags
            Catch ex As Exception
                Throw ex
            End Try
        End Function
        Public Function IsAssociated(ByVal ContactID As Int64) As Integer
            Dim dsRemSys As New DataSet
            Dim dtReturn As Integer = 0
            Dim strSQL As String

            Try

                dtReturn = oContactStructDB.DBGetFunction("dbo.udfConIsAssociated", ContactID)

                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Event Handlers"
        Private Sub evtValidationErr(ByVal MsgStr As String) Handles oContactDatum.evtValidationErr
            RaiseEvent evtContactStructErr(MsgStr)
        End Sub
#End Region
#End Region
    End Class
End Namespace
