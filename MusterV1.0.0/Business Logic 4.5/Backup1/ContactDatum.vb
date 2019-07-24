'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.ContactDatum
'   Provides the operations required to manipulate an ContactDatum object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0     KKM         03/28/05    Class definition.
'   1.1     MR          04/19/05    Added few Contact Properties.
'   1.2     MR          04/29/05    Added Ext1 and Ext2 Attributes
'
' Function          Description
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pContactDatum
#Region "Public Events"
        Public Event evtValidationErr(ByVal MsgStr As String)
#End Region
#Region "Private Member Variables"
        Private WithEvents oCOntactDatumInfo As MUSTER.Info.ContactDatumInfo
        Private oContactDatumDB As MUSTER.DataAccess.ContactDatumDB
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private colDatumCollection As MUSTER.Info.ContactDatumCollection
        Private oCOntactStructInfo As MUSTER.Info.ContactStructInfo
#End Region
#Region "Constructors"
        Public Sub New(Optional ByVal strDBConn As String = "", Optional ByRef MusterXCEP As MUSTER.Exceptions.MusterExceptions = Nothing, Optional ByRef ContactStructInfo As MUSTER.Info.ContactStructInfo = Nothing)
            If MusterXCEP Is Nothing Then
                MusterException = New MUSTER.Exceptions.MusterExceptions
            Else
                MusterException = MusterXCEP
            End If
            oCOntactDatumInfo = New MUSTER.Info.ContactDatumInfo
            oContactDatumDB = New MUSTER.DataAccess.ContactDatumDB
            colDatumCollection = New MUSTER.Info.ContactDatumCollection
        End Sub
#End Region
#Region "Exposed Attributes"

        Public Property companyName() As String
            Get
                Return oCOntactDatumInfo.companyName
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.companyName = Value
            End Set
        End Property

        Public ReadOnly Property ID() As Integer
            Get
                Return oCOntactDatumInfo.ID
            End Get
        End Property

        Public Property City() As String
            Get
                Return oCOntactDatumInfo.City
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.City = Value
            End Set
        End Property

        Public Property State() As String
            Get
                Return oCOntactDatumInfo.State
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.State = Value
            End Set
        End Property

        Public Property ZipCode() As String
            Get
                Return oCOntactDatumInfo.ZipCode
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.ZipCode = Value
            End Set
        End Property
       
        Public Property Deleted() As Boolean
            Get
                Return oCOntactDatumInfo.deleted
            End Get
            Set(ByVal Value As Boolean)
                oCOntactDatumInfo.deleted = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oCOntactDatumInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oCOntactDatumInfo.IsDirty = Value
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim oContactDatumInfoLocal As MUSTER.Info.ContactDatumInfo
                For Each oContactDatumInfoLocal In colDatumCollection.Values
                    If oContactDatumInfoLocal.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
            End Get
            Set(ByVal Value As Boolean)
            End Set
        End Property
        Public Property title() As String
            Get
                Return oCOntactDatumInfo.Title
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.Title = Value
            End Set
        End Property
        Public Property FirstName() As String
            Get
                Return oCOntactDatumInfo.FirstName
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.FirstName = Value
            End Set
        End Property
        Public Property MiddleName() As String
            Get
                Return oCOntactDatumInfo.MiddleName
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.MiddleName = Value
            End Set
        End Property
        Public Property Prefix() As String
            Get
                Return oCOntactDatumInfo.Prefix
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.Prefix = Value
            End Set
        End Property
        Public Property Lastname() As String
            Get
                Return oCOntactDatumInfo.LastName
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.LastName = Value
            End Set
        End Property
        Public Property AddressLine1() As String
            Get
                Return oCOntactDatumInfo.AddressLine1
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.AddressLine1 = Value
            End Set
        End Property
        Public Property AddressLine2() As String
            Get
                Return oCOntactDatumInfo.AddressLine2
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.AddressLine2 = Value
            End Set
        End Property
        Public Property Suffix() As String
            Get
                Return oCOntactDatumInfo.suffix
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.suffix = Value
            End Set
        End Property
        Public Property phone1() As String
            Get
                Return oCOntactDatumInfo.Phone1
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.Phone1 = Value
            End Set
        End Property
        Public Property Phone2() As String
            Get
                Return oCOntactDatumInfo.Phone2
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.Phone2 = Value
            End Set
        End Property
        Public Property Ext1() As String
            Get
                Return oCOntactDatumInfo.Ext1
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.Ext1 = Value
            End Set
        End Property
        Public Property Ext2() As String
            Get
                Return oCOntactDatumInfo.Ext2
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.Ext2 = Value
            End Set
        End Property
        Public Property fax() As String
            Get
                Return oCOntactDatumInfo.Fax
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.Fax = Value
            End Set
        End Property
        Public Property GeneralEmail() As String
            Get
                Return oCOntactDatumInfo.publicEmail
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.publicEmail = Value
            End Set
        End Property
        Public Property PersonalEmail() As String
            Get
                Return oCOntactDatumInfo.privateEmail
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.privateEmail = Value
            End Set
        End Property
        Public Property VendorNumber() As String
            Get
                Return oCOntactDatumInfo.VendorNumber
            End Get
            Set(ByVal Value As String)
                oCOntactDatumInfo.VendorNumber = Value
            End Set
        End Property
        Public ReadOnly Property fullName() As String
            Get
                Return oCOntactDatumInfo.fullName
            End Get
        End Property
        Public ReadOnly Property SortName() As String
            Get
                Return oCOntactDatumInfo.SortName
            End Get
        End Property
        Public ReadOnly Property PlainName() As String
            Get
                Return oCOntactDatumInfo.PlainName
            End Get
        End Property
        Public Property ContactCollection() As MUSTER.Info.ContactDatumCollection
            Get
                Return colDatumCollection
            End Get
            Set(ByVal Value As MUSTER.Info.ContactDatumCollection)
                colDatumCollection = Value
            End Set
        End Property
        Public Property contactDatumInfo() As MUSTER.Info.ContactDatumInfo
            Get
                Return oCOntactDatumInfo
            End Get
            Set(ByVal Value As MUSTER.Info.ContactDatumInfo)
                oCOntactDatumInfo = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub GetAll()
            Try
                colDatumCollection = oContactDatumDB.DBGetAll()
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function GetByID(ByVal nContactID As Integer, Optional ByVal entityID As Integer = 0, Optional ByVal moduleID As Integer = 0) As DataTable
            Dim dsSet As DataSet
            Try
                dsSet = oContactDatumDB.DBGetByID(nContactID, entityID, moduleID)
                Return dsSet.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Sub GetAllByEntity(ByVal EntityID As Int64, ByVal ModuleID As Int64)
            Try
                colDatumCollection = oContactDatumDB.DBGetAll(EntityID, ModuleID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Retrieve(ByVal contactID As Integer, Optional ByVal showDeleted As Boolean = False, Optional ByRef ContactStructInfo As MUSTER.Info.ContactStructInfo = Nothing) As MUSTER.Info.ContactDatumInfo
            Dim bolDataAged As Boolean = False
            Try
                oCOntactStructInfo = ContactStructInfo
                If Not (oCOntactDatumInfo.deleted Or oCOntactDatumInfo.ID = 0) Then
                End If
                Dim oContactDatumInfoLocal As MUSTER.Info.ContactDatumInfo

                If contactID = 0 Then
                    Add(0)
                Else
                    ' get by contact id
                    oCOntactDatumInfo = colDatumCollection.Item(contactID.ToString)
                    ' Check for Aged Data here.
                    If Not (oCOntactDatumInfo Is Nothing) Then
                        If oCOntactDatumInfo.IsAgedData = True And oCOntactDatumInfo.IsDirty = False Then
                            bolDataAged = True
                            colDatumCollection.Remove(oCOntactDatumInfo)
                        End If
                    End If
                    If oCOntactDatumInfo Is Nothing Or bolDataAged Then
                        Add(contactID, showDeleted)
                    End If
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Return oCOntactDatumInfo
        End Function
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolValidated As Boolean = False, Optional ByVal bolDelete As Boolean = False, Optional ByVal bolAdddressDirty As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If bolAdddressDirty Then
                    '  Return True
                End If
                If Not bolValidated And Not oCOntactDatumInfo.deleted And Not bolDelete Then
                    'If Not Me.ValidateData() Then
                    '    Return False
                    'End If
                End If
                If Not (oCOntactDatumInfo.ID < 0 And oCOntactDatumInfo.deleted) Then
                    oldID = oCOntactDatumInfo.ID

                    oContactDatumDB.Put(oCOntactDatumInfo, moduleID, staffID, returnVal, (bolAdddressDirty))

                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If
                    If Not bolValidated Then
                        If oldID <> oCOntactDatumInfo.ID Then
                            colDatumCollection.ChangeKey(oldID, oCOntactDatumInfo.ID)
                        End If
                    End If
                    oCOntactDatumInfo.Archive()
                    oCOntactDatumInfo.IsDirty = False
                End If
                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        'Function ValidateData() As Boolean
        '    Try
        '        Dim errStr As String = ""
        '        Dim validateSuccess As Boolean = True

        '        If (oCOntactDatumInfo.FirstName = String.Empty Or oCOntactDatumInfo.LastName = String.Empty) And oCOntactDatumInfo.IsPerson = True Then
        '            errStr += "Please check the FIRST NAME and the LAST NAME (required)" + vbCrLf
        '            validateSuccess = False
        '        End If
        '        If oCOntactDatumInfo.companyName = String.Empty And oCOntactDatumInfo.IsPerson = False Then
        '            errStr += "Please Enter the COMPANY NAME(required)" + vbCrLf
        '            validateSuccess = False
        '        End If
        '        If (oCOntactDatumInfo.AddressLine1 <> String.Empty Or oCOntactDatumInfo.ZipCode <> "" Or oCOntactDatumInfo.City <> String.Empty) Or Not oCOntactDatumInfo.IsPerson Then
        '            If oCOntactDatumInfo.AddressLine1 = String.Empty Then
        '                errStr += "Please check the ADDRESS(required)" + vbCrLf
        '                validateSuccess = False
        '            End If
        '            If Not ValidatePhone(oCOntactDatumInfo.ZipCode) And oCOntactDatumInfo.ZipCode = "_____-____" Then
        '                errStr += "Please check the ZIP code(required)" + vbCrLf
        '                validateSuccess = False
        '            End If
        '            If oCOntactDatumInfo.City = String.Empty Or oCOntactDatumInfo.State = String.Empty Then
        '                errStr += "Please check the CITY and STATE(required)" + vbCrLf
        '                validateSuccess = False
        '            End If
        '        End If
        '        If oCOntactDatumInfo.Phone1 <> "(___)___-____" And Not ValidatePhone(oCOntactDatumInfo.Phone1) Then
        '            errStr += "Please check the PHONE1" + vbCrLf
        '            validateSuccess = False
        '        End If
        '        If oCOntactDatumInfo.Phone2 <> "(___)___-____" And Not ValidatePhone(oCOntactDatumInfo.Phone2) Then
        '            errStr += "Please check the PHONE2" + vbCrLf
        '            validateSuccess = False
        '        End If

        '        If oCOntactDatumInfo.Cell <> "(___)___-____" Then
        '            If Not ValidatePhone(oCOntactDatumInfo.Cell) Then
        '                errStr += "Please check the CELL number" + vbCrLf
        '                validateSuccess = False
        '            End If
        '        End If
        '        If oCOntactDatumInfo.Fax <> "(___)___-____" Then
        '            If Not ValidatePhone(oCOntactDatumInfo.Fax) Then
        '                errStr += "Please check the FAX number" + vbCrLf
        '                validateSuccess = False
        '            End If
        '        End If
        '        If errStr.Length > 0 And Not validateSuccess Then
        '            RaiseEvent evtValidationErr(errStr)
        '        End If
        '        Return validateSuccess
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
        Private Function ValidatePhone(ByVal strPhone As String) As Boolean
            Try
                Dim strRegex As String = "(\(\d\d\d\))?\s*(\d\d\d)\s*[\-]?\s*(\d\d\d\d)"
                '"^\(?\d{3}\)?\s|-\d{3}-\d{4}$" -  matches (555) 555-5555, or 555-555-5555
                Dim rx As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(strRegex)
                If rx.IsMatch(strPhone) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Public Sub Add(ByVal id As Integer, Optional ByVal showDeleted As Boolean = False)
            Try
                oCOntactDatumInfo = oContactDatumDB.DBGetInfoByID(id, showDeleted)
                If oCOntactDatumInfo.ID = 0 Then
                    oCOntactDatumInfo.ID = nID
                    nID -= 1
                End If
                colDatumCollection.Add(oCOntactDatumInfo)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Sub Add(ByRef oContactDatum As MUSTER.Info.ContactDatumInfo)
            Try
                oCOntactDatumInfo = oContactDatum
                colDatumCollection.Add(oCOntactDatumInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Remove(ByVal id As Integer)
            Dim oContactDatumInfoLocal As MUSTER.Info.ContactDatumInfo
            Try
                oCOntactDatumInfo = colDatumCollection.Item(id)
                If Not (oContactDatumInfoLocal Is Nothing) Then
                    colDatumCollection.Remove(oContactDatumInfoLocal)
                    Exit Sub
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
            Throw New Exception("Closure Event " & id.ToString & " is not in the collection of Closure Events.")
        End Sub
        Public Sub Remove(ByVal oContactDatumInf As MUSTER.Info.ContactDatumInfo)
            Try
                colDatumCollection.Remove(oContactDatumInf)
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Closure Event " & oContactDatumInf.ID & " is not in the collection of Closure Events.")
        End Sub
        Public Sub Flush()
            Try
            Catch ex As Exception
            End Try
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oCOntactDatumInfo = New MUSTER.Info.ContactDatumInfo
        End Sub
        Public Sub Reset()
            oCOntactDatumInfo.Reset()
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colDatumCollection.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 Then
                If colIndex + direction <= nArr.GetUpperBound(0) Then
                    Return colDatumCollection.Item(nArr.GetValue(colIndex + direction)).ID.ToString
                Else
                    Return colDatumCollection.Item(nArr.GetValue(0)).ID.ToString
                End If
            Else
                Return colDatumCollection.Item(nArr.GetValue(nArr.GetUpperBound(0))).ID.ToString
            End If
        End Function
        Public Function Clone(Optional ByVal strTarget As String = Nothing) As MUSTER.Info.ContactDatumInfo
            If Not strTarget Is Nothing Then
                Return oCOntactDatumInfo
            ElseIf strTarget = "In Module" Then

            End If
        End Function
#End Region
#End Region
#Region "Miscellaneous operations"
        Public Function GetCompanyAddress(ByVal nContactID As Integer) As DataTable
            Dim dsSet As DataSet
            Try
                dsSet = oContactDatumDB.DBGetDS("spCONContactGetCompanyAddress", nContactID)
                Return dsSet.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function EntityTable() As DataTable
            Dim oContactDatumInfoLocal As New MUSTER.Info.ContactDatumInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("ContactID")
                tbEntityTable.Columns.Add("Contact")
                tbEntityTable.Columns.Add("Address")
                tbEntityTable.Columns.Add("City")
                tbEntityTable.Columns.Add("State")
                tbEntityTable.Columns.Add("ZipCode")
                tbEntityTable.Columns.Add("IsPerson")
                For Each oContactDatumInfoLocal In colDatumCollection.Values
                    dr = tbEntityTable.NewRow()
                    dr("ContactID") = oContactDatumInfoLocal.ID
                    dr("City") = oContactDatumInfoLocal.City
                    dr("State") = oContactDatumInfoLocal.State
                    dr("ZipCode") = oContactDatumInfoLocal.ZipCode
                    dr("IsPerson") = oContactDatumInfoLocal.IsPerson
                    tbEntityTable.Rows.Add(dr)
                Next
                Return tbEntityTable
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function GetContactLastName(ByVal nContactID As Integer) As DataSet
            Dim dsSet As DataSet
            Try
                dsSet = oContactDatumDB.DBGetDS("spCONContactGetLastname", nContactID)
                'dsSet = oContactDatumDB.DBGetDS("Select Title,Last_Name,Suffix from tblcon_contacts where ContactID=" + nContactID.ToString)
                Return dsSet
            Catch ex As Exception

            End Try
        End Function
#End Region
    End Class
End Namespace
