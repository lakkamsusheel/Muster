'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Address
'   Provides the info and collection objects to the client for manipulating
'   an ComAddressInfo object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         MR      5/24/05    Original class definition.
'
' Function                          Description
' Retrieve(ID)       Sets the internal ComAddressInfo to the ComAddressInfo matching the supplied key.  
' GetAddress(ID)     Returns the Address requested by the int arg ID
' GetDataSet(strSQL) Returns a DataSet from the repository using strSQL as supplied.
' GetAddressAll()    Returns an ComAddressCollection with all ComAddressInfo objects
' Add(ID)            Adds the Address identified by arg ID to the 
'                           internal ComAddressCollection
' Add(ComAddressInfo)   Adds the ComAddressInfo passed as the argument to the internal 
'                           ComAddressCollection
' Remove(ID)         Removes the ComAddressInfo identified by arg ID from the internal 
'                           ComAddressCollection
' colIsDirty()       Returns a boolean indicating whether any of the ComAddressInfo
'                    objects in the ComAddressCollection has been modified since the
'                    last time it was retrieved from/saved to the repository.
' Flush()            Marshalls all modified/added ComAddressInfo objects in the 
'                        ComAddressCollection to the repository.
' Save()             Marshalls the internal ComAddressInfo object to the repository.
' Save(ComAddressInfo)  Marshalls the ComAddressInfo object to the repository as supplied in Args.
' Clear()
' Reset()
' AddressTable()     Returns a datatable containing all columns for the ComAddressInfo
'                           objects in the internal ComAddressCollection.
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pComAddress
#Region "Public Events"
        Public Event evtAddressErr(ByVal MsgStr As String)
        Public Event evtAddressChanged(ByVal bolValue As Boolean)
        Public Event evtAddressesChanged(ByVal bolValue As Boolean)
        Public Event CitiesChanged(ByVal dsCities As DataSet)
        Public Event ZipChanged(ByVal dsZip As DataSet)
        Public Event StateChanged(ByVal dsState As DataSet)
        Public Event FipsChanged(ByVal strFIPS As String)

#End Region
#Region "Private Member Variables"
        Private WithEvents colComAddresses As MUSTER.Info.ComAddressCollection
        Private WithEvents oComAddressInfo As MUSTER.Info.ComAddressInfo
        Private oComAddressDB As New MUSTER.DataAccess.ComAddressDB
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private nNewIndex As Integer = 0
        Private blnShowDeleted As Boolean = False
        Private nID As Int64 = -1
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private strFIPS As String
        Private bolSuccess As Boolean = False
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Address").ID

#End Region
#Region "Constructors"
        Public Sub New()
            oComAddressInfo = New MUSTER.Info.ComAddressInfo
            colComAddresses = New MUSTER.Info.ComAddressCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ColCompanyAddresses() As MUSTER.Info.ComAddressCollection
            Get
                Return colComAddresses
            End Get
            Set(ByVal Value As MUSTER.Info.ComAddressCollection)
                colComAddresses = Value
            End Set
        End Property

        Public Property AddressId() As Integer
            Get
                Return oComAddressInfo.AddressId
            End Get

            Set(ByVal value As Integer)
                oComAddressInfo.AddressId = Integer.Parse(value)
            End Set
        End Property
        Public Property CompanyId() As Integer
            Get
                Return oComAddressInfo.CompanyId
            End Get

            Set(ByVal value As Integer)
                oComAddressInfo.CompanyId = Integer.Parse(value)
            End Set
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return oComAddressInfo.LicenseeID
            End Get

            Set(ByVal value As Integer)
                oComAddressInfo.LicenseeID = Integer.Parse(value)
            End Set
        End Property
        Public Property ProviderID() As Integer
            Get
                Return oComAddressInfo.ProviderID
            End Get

            Set(ByVal value As Integer)
                oComAddressInfo.ProviderID = Integer.Parse(value)
            End Set
        End Property
        Public Property AddressLine1() As String
            Get
                Return oComAddressInfo.AddressLine1
            End Get

            Set(ByVal value As String)
                oComAddressInfo.AddressLine1 = value
            End Set
        End Property
        Public Property AddressLine2() As String
            Get
                Return oComAddressInfo.AddressLine2
            End Get

            Set(ByVal value As String)
                oComAddressInfo.AddressLine2 = value
            End Set
        End Property
        Public Property City() As String
            Get
                Return oComAddressInfo.City
            End Get

            Set(ByVal value As String)
                oComAddressInfo.City = value
            End Set
        End Property
        Public Property State() As String
            Get
                Return oComAddressInfo.State
            End Get

            Set(ByVal value As String)
                oComAddressInfo.State = value
            End Set
        End Property
        Public Property Zip() As String
            Get
                Return oComAddressInfo.Zip
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Zip = value
            End Set
        End Property
        Public Property FIPSCode() As String
            Get
                Return oComAddressInfo.FIPSCode
            End Get

            Set(ByVal value As String)
                oComAddressInfo.FIPSCode = value
            End Set
        End Property
        Public Property Phone1() As String
            Get
                Return oComAddressInfo.Phone1
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Phone1 = value
            End Set
        End Property
        Public Property Phone2() As String
            Get
                Return oComAddressInfo.Phone2
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Phone2 = value
            End Set
        End Property
        Public Property Ext1() As String
            Get
                Return oComAddressInfo.Ext1
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Ext1 = value
            End Set
        End Property
        Public Property Ext2() As String
            Get
                Return oComAddressInfo.Ext2
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Ext2 = value
            End Set
        End Property
        Public Property Phone1Comment() As String
            Get
                Return oComAddressInfo.Phone1Comment
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Phone1Comment = value
            End Set
        End Property
        Public Property Phone2Comment() As String
            Get
                Return oComAddressInfo.Phone2Comment
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Phone2Comment = value
            End Set
        End Property
        Public Property Cell() As String
            Get
                Return oComAddressInfo.Cell
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Cell = value
            End Set
        End Property
        Public Property Fax() As String
            Get
                Return oComAddressInfo.Fax
            End Get

            Set(ByVal value As String)
                oComAddressInfo.Fax = value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oComAddressInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oComAddressInfo.Deleted = value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                If oComAddressInfo.IsDirty Then
                    Return True
                Else
                    Return False
                End If
            End Get

            Set(ByVal value As Boolean)
                oComAddressInfo.IsDirty = value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xAddressInfo As MUSTER.Info.ComAddressInfo

                For Each xAddressInfo In colComAddresses.Values
                    If xAddressInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property
        Public Property ShowDeleted() As Boolean
            Get
                Return blnShowDeleted
            End Get

            Set(ByVal value As Boolean)
                blnShowDeleted = value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oComAddressInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oComAddressInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oComAddressInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oComAddressInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oComAddressInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oComAddressInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByVal dr As DataRow)
            Try
                oComAddressInfo = New MUSTER.Info.ComAddressInfo(dr)
                colComAddresses.Add(oComAddressInfo)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False)
            Dim oldID As Integer
            Try
                If Not bolValidated Then
                    'If Not Me.ValidateData() Then
                    'End If
                End If
                If Not (oComAddressInfo.AddressId < 0 And oComAddressInfo.Deleted) Then
                    oldID = oComAddressInfo.AddressId
                    oComAddressDB.PutAddress(oComAddressInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oldID <> oComAddressInfo.AddressId Then
                            colComAddresses.ChangeKey(oldID, oComAddressInfo.AddressId)
                        End If
                    End If
                    oComAddressInfo.Archive()
                    oComAddressInfo.IsDirty = False
                End If
                RaiseEvent evtAddressChanged(oComAddressInfo.IsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Save(ByRef oAddress As MUSTER.Info.ComAddressInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim oldID As Integer
            Try
                oldID = oAddress.AddressId
                oComAddressDB.PutAddress(oAddress, moduleID, staffID, returnVal)
                If oldID <> oAddress.AddressId Then
                    colComAddresses.ChangeKey(oldID, oAddress.AddressId)
                End If
                oAddress.Archive()
                oAddress.IsDirty = False
                RaiseEvent evtAddressChanged(oAddress.IsDirty)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Obtains and returns an Address as called for by ID
        Public Function Retrieve(ByVal ID As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.ComAddressInfo

            Dim oComAddressInfoLocal As MUSTER.Info.ComAddressInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oComAddressInfoLocal In colComAddresses.Values
                    If oComAddressInfoLocal.AddressId = ID Then
                        If oComAddressInfoLocal.IsDirty = False And oComAddressInfoLocal.IsAgedData = True Then
                            bolDataAged = True
                        Else
                            oComAddressInfo = oComAddressInfoLocal
                            Return oComAddressInfo
                        End If
                    End If
                Next
                If bolDataAged = True Then
                    colComAddresses.Remove(oComAddressInfoLocal)
                End If
                oComAddressInfo = oComAddressDB.DBGetByID(ID)
                If oComAddressInfo.AddressId = 0 Then
                    oComAddressInfo.AddressId = nID
                    nID -= 1
                End If
                colComAddresses.Add(oComAddressInfo)
                Return oComAddressInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(Optional ByVal nProviderID As Int32 = 0) As MUSTER.Info.ComAddressInfo

            Dim oComAddressInfoLocal As MUSTER.Info.ComAddressInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oComAddressInfoLocal In colComAddresses.Values
                    If oComAddressInfoLocal.ProviderID = nProviderID Then
                        oComAddressInfo = oComAddressInfoLocal
                        Return oComAddressInfo
                    End If
                Next

                oComAddressInfo = oComAddressDB.DBGetByProviderID(nProviderID)
                If oComAddressInfo.AddressId = 0 Then
                    oComAddressInfo.AddressId = nID
                    nID -= 1
                End If
                colComAddresses.Add(oComAddressInfo)
                Return oComAddressInfo
                Return oComAddressInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetAddress(ByVal ID As Int64) As MUSTER.Info.ComAddressInfo
            Dim oAddressInfoLocal As MUSTER.Info.ComAddressInfo
            Try
                For Each oAddressInfoLocal In colComAddresses.Values
                    If oAddressInfoLocal.AddressId = ID Then
                        oComAddressInfo = oAddressInfoLocal
                        Return oComAddressInfo
                    End If
                Next
                oComAddressInfo = oComAddressDB.DBGetByID(ID)
                colComAddresses.Add(oComAddressInfo)
                Return oComAddressInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCompanyAddress(ByVal nCompanyID As Integer) As DataSet
            Try
                Return oComAddressDB.DBGetCompanyAddress(nCompanyID, False)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetAddressByType(ByVal nAddressID As Integer, Optional ByVal nCompanyID As Integer = 0, Optional ByVal nLicenseeID As Integer = 0, Optional ByVal nProviderID As Integer = 0) As MUSTER.Info.ComAddressInfo
            Try
                oComAddressInfo = oComAddressDB.DBGetByTypeID(nAddressID, nCompanyID, nLicenseeID, nProviderID)
                Return oComAddressInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Function GetDataSet(ByVal strSQL As String) As DataSet
            Try
                Dim ds As DataSet
                ds = oComAddressDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Validates Data according to DDD Specifications
        'Public Function ValidateData() As Boolean
        '    Try
        '        Dim errStr As String = ""
        '        Dim validateSuccess As Boolean = True

        '        If oComAddressInfo.AddressId <> 0 Then
        '            If oComAddressInfo.AddressLine1 <> String.Empty Then
        '                If oComAddressInfo.City <> String.Empty Then
        '                    If oComAddressInfo.State <> String.Empty Then
        '                        If oComAddressInfo.Zip <> String.Empty Then
        '                            validateSuccess = True
        '                        Else
        '                            errStr += "Zip cannot be empty" + vbCrLf
        '                            validateSuccess = False
        '                        End If
        '                    Else
        '                        errStr += "State cannot be empty" + vbCrLf
        '                        validateSuccess = False
        '                    End If
        '                Else
        '                    errStr += "City cannot be empty" + vbCrLf
        '                    validateSuccess = False
        '                End If
        '            Else
        '                errStr += "AddressLine1 cannot be empty" + vbCrLf
        '                validateSuccess = False
        '            End If
        '        End If
        '        If errStr.Length > 0 Or Not validateSuccess Then
        '            RaiseEvent evtAddressErr(errStr)
        '        End If
        '        Return validateSuccess
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
#End Region
#Region "Collection Operations"
        Function GetAddressAll(Optional ByVal nAddressID As Integer = 0, Optional ByVal nCompanyID As Integer = 0, Optional ByVal nLicenseeID As Integer = 0, Optional ByVal nProviderID As Integer = 0, Optional ByVal blnShowDeleted As Boolean = False) As MUSTER.Info.ComAddressCollection
            colComAddresses.Clear()
            Try
                colComAddresses = oComAddressDB.GetAllInfo(nAddressID, nCompanyID, nLicenseeID, nProviderID, blnShowDeleted)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Return colComAddresses
        End Function
        Public Sub Add(ByVal ID As Int64, Optional ByVal showDeleted As Boolean = False)
            Try
                oComAddressInfo = oComAddressDB.DBGetByID(ID, showDeleted)
                If oComAddressInfo.AddressId = 0 Then
                    oComAddressInfo.AddressId = nID
                    nID -= 1
                End If
                colComAddresses.Add(oComAddressInfo)
            Catch ex As Exception
                Throw ex
            End Try

        End Sub
        Public Function Add(ByRef oAddress As MUSTER.Info.ComAddressInfo) As Boolean
            Try

                oComAddressInfo = oAddress
                'If ValidateData() Then
                If oComAddressInfo.AddressId = 0 Then
                    oComAddressInfo.AddressId = nID
                    nID -= 1
                End If
                colComAddresses.Add(oComAddressInfo)
                'Return True
                'Else
                'Return False
                'End If

            Catch ex As Exception
                Throw ex
            End Try
        End Function
        Public Function Add() As Boolean
            Try
                'If ValidateData() Then
                If oComAddressInfo.AddressId = 0 Then
                    oComAddressInfo.AddressId = nID
                    nID -= 1
                End If
                colComAddresses.Add(oComAddressInfo)
                'Return True
                'Else
                'Return False
                'End If
            Catch ex As Exception
                Throw ex
            End Try
        End Function
        Public Sub Remove(ByVal ID As Int64)
            Dim myIndex As Int16 = 1
            Dim oAddressInfoLocal As MUSTER.Info.ComAddressInfo
            Try
                oAddressInfoLocal = colComAddresses.Item(ID)
                If Not (oAddressInfoLocal Is Nothing) Then
                    colComAddresses.Remove(oAddressInfoLocal)
                    Exit Sub
                End If

            Catch ex As Exception
                Throw ex
            End Try

            Throw New Exception("Com Address " & ID.ToString & " is not in the collection of Company Addresses.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByRef pCompany As MUSTER.BusinessLogic.pCompany = Nothing, Optional ByRef pComLicenseeAssoc As MUSTER.BusinessLogic.pCompanyLicensee = Nothing)
            Try
                Dim IDs As New Collection
                Dim index As Integer
                Dim xComAddressInfo As MUSTER.Info.ComAddressInfo
                For Each xComAddressInfo In colComAddresses.Values
                    If xComAddressInfo.IsDirty Then
                        If pCompany.ID > 0 Then
                            xComAddressInfo.CompanyId = pCompany.ID
                        End If
                        oComAddressInfo = xComAddressInfo
                        IDs.Add(oComAddressInfo.AddressId)
                        Me.Save(moduleID, staffID, returnVal, True)
                    End If
                Next

                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        xComAddressInfo = colComAddresses.Item(colKey)
                        colComAddresses.ChangeKey(colKey, xComAddressInfo.AddressId)

                        'Update the AddressID in Company Licensee Association Collection
                        If Not pComLicenseeAssoc Is Nothing Then
                            Dim xComLicInfo As MUSTER.Info.CompanyLicenseeInfo
                            'xComLicInfo = pComLicenseeAssoc.ComLicCollection.Item(colKey)
                            For Each xComLicInfo In pComLicenseeAssoc.ComLicCollection.Values
                                If xComLicInfo.ComLicAddressID = CType(colKey, Integer) Then
                                    xComLicInfo.ComLicAddressID = xComAddressInfo.AddressId
                                End If
                            Next
                        End If

                        If Not pCompany Is Nothing Then
                            If pCompany.PRO_ENGIN_ADD_ID = CType(colKey, Integer) Then
                                pCompany.PRO_ENGIN_ADD_ID = xComAddressInfo.AddressId
                            End If

                            If pCompany.PRO_GEOLO_ADD_ID = CType(colKey, Integer) Then
                                pCompany.PRO_GEOLO_ADD_ID = xComAddressInfo.AddressId
                            End If

                        End If



                    Next
                End If




            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oComAddressInfo = New MUSTER.Info.ComAddressInfo
        End Sub
        Public Sub Reset()
            oComAddressInfo.Reset()
        End Sub
#End Region
#Region "LookUp Operations"

        Public Function PopulateAddressDetails(Optional ByVal strCity As String = "", Optional ByVal strState As String = "", Optional ByVal strZip As String = "")

            Dim strWhere As String = String.Empty
            Dim strMessage As String = ""
            Dim dsTemp As DataSet
            Try
                bolSuccess = True
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
                    If dsTemp.Tables(0).Rows.Count > 0 Then
                        RaiseEvent StateChanged(dsTemp)
                        dsTemp.Tables.Clear()
                    Else
                        strMessage += vbTab + "State cannot be determined" + vbCrLf
                    End If
                End If

                If strState = String.Empty And strCity = String.Empty And strZip = String.Empty Then
                    strWhere = " WHERE UPPER(STATE) = 'MS'"
                End If

                If strCity = String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT CITY from tblSYS_ZIPCODES " & strWhere & " ORDER BY CITY")
                    If dsTemp.Tables(0).Rows.Count > 0 Then
                        RaiseEvent CitiesChanged(dsTemp)
                        dsTemp.Tables.Clear()
                    Else
                        strMessage += vbTab + "City cannot be determined" + vbCrLf
                    End If
                End If

                If strZip = String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT ZIP  from tblSYS_ZIPCODES " & strWhere & "  ORDER BY ZIP")
                    If dsTemp.Tables(0).Rows.Count > 0 Then
                        RaiseEvent ZipChanged(dsTemp)
                        dsTemp.Tables.Clear()
                    Else
                        strMessage += vbTab + "Zip cannot be determined" + vbCrLf
                    End If
                End If

                If strState <> String.Empty And strCity <> String.Empty And strZip <> String.Empty Then
                    dsTemp = PopulateAddressLookUp("Select DISTINCT FIPS from tblSYS_ZIPCODES " & strWhere & " ORDER BY FIPS")
                    If dsTemp.Tables(0).Rows.Count > 0 Then
                        strFIPS = IIf(IsDBNull(dsTemp.Tables(0).Rows(0).Item("FIPS")), "", dsTemp.Tables(0).Rows(0).Item("FIPS"))
                        RaiseEvent FipsChanged(strFIPS)
                        dsTemp.Tables.Clear()
                    Else
                        strMessage += vbTab + "FIPS cannot be determined" + vbCrLf
                    End If

                End If

                If strMessage.Length > 0 Then
                    bolSuccess = False
                    Throw New Exception(strMessage + "Please Enter Valid Data.")
                End If

            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Function PopulateAddressLookUp(ByVal strSQL As String) As DataSet
            Dim dsReturn As New DataSet

            Try
                dsReturn = oComAddressDB.DBGetDS(strSQL)
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the addresses in the collection
        Public Function AddressTable() As DataTable

            Dim oAddressInfoLocal As MUSTER.Info.ComAddressInfo
            Dim dr As DataRow
            Dim tbAddressTable As New DataTable

            Try
                tbAddressTable.Columns.Add("Address 1", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("Address 2", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("City", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("State", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("Zip", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("Phone 1", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("Ext1", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("Phone 2", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("Ext2", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("Fax", Type.GetType("System.String"))
                tbAddressTable.Columns.Add("ADDRESS_ID", Type.GetType("System.Int64"))
                tbAddressTable.Columns.Add("Company_ID", Type.GetType("System.Int64"))
                tbAddressTable.Columns.Add("Licensee_ID", Type.GetType("System.Int64"))
                tbAddressTable.Columns.Add("Provider_ID", Type.GetType("System.Int64"))

                For Each oAddressInfoLocal In colComAddresses.Values
                    dr = tbAddressTable.NewRow()
                    dr("Address 1") = oAddressInfoLocal.AddressLine1
                    dr("Address 2") = oAddressInfoLocal.AddressLine2
                    dr("City") = oAddressInfoLocal.City
                    dr("State") = oAddressInfoLocal.State
                    dr("Zip") = oAddressInfoLocal.Zip
                    dr("Phone 1") = oAddressInfoLocal.Phone1
                    dr("Ext1") = oAddressInfoLocal.Ext1
                    dr("Phone 2") = oAddressInfoLocal.Phone2
                    dr("Ext2") = oAddressInfoLocal.Ext2
                    dr("Fax") = oAddressInfoLocal.Fax
                    dr("ADDRESS_ID") = oAddressInfoLocal.AddressId
                    dr("Company_ID") = oAddressInfoLocal.CompanyId
                    dr("Licensee_ID") = oAddressInfoLocal.LicenseeID
                    dr("Provider_ID") = oAddressInfoLocal.ProviderID
                    tbAddressTable.Rows.Add(dr)
                Next
                Return tbAddressTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
#Region "Event Handlers"
        Private Sub AddressesChanged(ByVal strSrc As String) Handles colComAddresses.AddressColChanged
            RaiseEvent evtAddressesChanged(Me.colIsDirty)
        End Sub
        Private Sub AddressChanged(ByVal bolValue As Boolean) Handles oComAddressInfo.AddressInfoChanged
            RaiseEvent evtAddressChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
