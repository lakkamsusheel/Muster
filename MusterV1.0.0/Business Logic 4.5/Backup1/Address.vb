'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Address
'   Provides the info and collection objects to the client for manipulating
'   an AddressInfo object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         KJ      12/07/04    Original class definition.
'   1.1         KJ      12/14/04    Changed from Temp Collection to single Collection
'                                   The logic might be convoluted but it is working. Need to revisit to modify this.
'   1.2         AN      12/15/04    Made Changes to GetAddress 
'   1.3         KJ      12/22/04    Made Changes to Descriptions in the Header
'   1.4         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.5         KJ      01/03/05    Changed the exposed attributes properties to have collection.
'   1.6         KJ      01/04/05    Changed the Retrieve Function. Commented the GetAddress Function as the same functionality is in Retrieve
'                                   Changed the Add Function to have addition of new object using -1
'   1.7         MNR     01/13/05    Modified Retrieve function to handle hirearchy, Added Events
'   1.8         MNR     01/14/05    Added ValidateData()
'   1.9         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   2.0         JVC2    02/02/05    Added EntityTypeID to private members and initialize to "Persona" type.'
'                                       Also added new attribute to expose the Type ID.
'   2.1         AB      02/17/05    Added DataAge check to the Retrieve function
'   2.2         MNR     03/15/05    Added Load Sub
'   2.3         MNR     03/16/05    Removed strSrc from events
'   2.4         MNR     07/25/05    Added county property
'   2.5  Thomas Franey  12/02/09    Added Physical Town property  
'
' Function                          Description
' Retrieve(ID)       Sets the internal TankInfo to the TankInfo matching the supplied key.  
' GetAddress(ID)     Returns the Address requested by the int arg ID
' GetDataSet(strSQL) Returns a DataSet from the repository using strSQL as supplied.
' GetAddressAll()    Returns an AddressCollection with all AddressInfo objects
' Add(ID)            Adds the Address identified by arg ID to the 
'                           internal AddressCollection
' Add(AddressInfo)   Adds the AddressInfo passed as the argument to the internal 
'                           AddressCollection
' Remove(ID)         Removes the AddressInfo identified by arg ID from the internal 
'                           AddressCollection
' colIsDirty()       Returns a boolean indicating whether any of the AddressInfo
'                    objects in the AddressCollection has been modified since the
'                    last time it was retrieved from/saved to the repository.
' Flush()            Marshalls all modified/added AddressInfo objects in the 
'                        AddressCollection to the repository.
' Save()             Marshalls the internal AddressInfo object to the repository.
' Save(AddressInfo)  Marshalls the AddressInfo object to the repository as supplied in Args.
' Clear()
' Reset()
' AddressTable()     Returns a datatable containing all columns for the AddressInfo
'                           objects in the internal AddressCollection.
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pAddress
#Region "Public Events"
        Public Event evtAddressErr(ByVal MsgStr As String)
        Public Event evtAddressChanged(ByVal bolValue As Boolean)
        Public Event evtAddressesChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents colAddresses As MUSTER.Info.AddressCollection
        Private WithEvents oAddressInfo As MUSTER.Info.AddressInfo
        Private oAddressDB As New MUSTER.DataAccess.AddressDB
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private nNewIndex As Integer = 0
        Private blnShowDeleted As Boolean = False
        Private nID As Int64 = -1
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Address").ID
#End Region
#Region "Constructors"
        Public Sub New()
            oAddressInfo = New MUSTER.Info.AddressInfo
            colAddresses = New MUSTER.Info.AddressCollection
        End Sub
#End Region
#Region "Exposed Attributes"

        Public Property AddressId() As Integer
            Get
                Return oAddressInfo.AddressId
            End Get

            Set(ByVal value As Integer)
                oAddressInfo.AddressId = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property AddressTypeId() As Integer
            Get
                Return oAddressInfo.AddressTypeId
            End Get

            Set(ByVal value As Integer)
                oAddressInfo.AddressTypeId = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property EntityType() As Integer
            Get
                Return oAddressInfo.EntityType
            End Get

            Set(ByVal value As Integer)
                oAddressInfo.EntityType = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property AddressLine1() As String
            Get
                Return oAddressInfo.AddressLine1
            End Get

            Set(ByVal value As String)
                oAddressInfo.AddressLine1 = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property


        Public Property PhysicalTown() As String
            Get
                Return oAddressInfo.PhsycalTown
            End Get

            Set(ByVal value As String)
                oAddressInfo.PhsycalTown = value
            End Set
        End Property

        Public Property AddressLine1ForEnsite() As String
            Get
                Return oAddressInfo.AddressLine1ForEnsite
            End Get

            Set(ByVal value As String)
                oAddressInfo.AddressLine1ForEnsite = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property AddressLine2() As String
            Get
                Return oAddressInfo.AddressLine2
            End Get

            Set(ByVal value As String)
                oAddressInfo.AddressLine2 = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property AddressLine2ForEnsite() As String
            Get
                Return oAddressInfo.AddressLine2ForEnsite
            End Get

            Set(ByVal value As String)
                oAddressInfo.AddressLine2ForEnsite = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property City() As String
            Get
                Return oAddressInfo.City
            End Get

            Set(ByVal value As String)
                oAddressInfo.City = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property State() As String
            Get
                Return oAddressInfo.State
            End Get

            Set(ByVal value As String)
                oAddressInfo.State = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property Zip() As String
            Get
                Return oAddressInfo.Zip
            End Get

            Set(ByVal value As String)
                oAddressInfo.Zip = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property FIPSCode() As String
            Get
                Return oAddressInfo.FIPSCode
            End Get

            Set(ByVal value As String)
                oAddressInfo.FIPSCode = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property County() As String
            Get
                Return oAddressInfo.County
            End Get
            Set(ByVal Value As String)
                oAddressInfo.County = Value
            End Set
        End Property

        Public WriteOnly Property CountyFirstTime() As String
            Set(ByVal Value As String)
                oAddressInfo.CountyFirstTime = Value
            End Set
        End Property

        Public Property StartDate() As DateTime
            Get
                Return oAddressInfo.StartDate
            End Get

            Set(ByVal value As DateTime)
                oAddressInfo.StartDate = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property EndDate() As DateTime
            Get
                Return oAddressInfo.EndDate
            End Get

            Set(ByVal value As DateTime)
                oAddressInfo.EndDate = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property Deleted() As Boolean
            Get
                Return oAddressInfo.Deleted
            End Get

            Set(ByVal value As Boolean)
                oAddressInfo.Deleted = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                If oAddressInfo.IsDirty Then
                    Return True
                Else
                    Return False
                End If
            End Get

            Set(ByVal value As Boolean)
                oAddressInfo.IsDirty = value
                'colAddresses(oAddressInfo.AddressId) = oAddressInfo
            End Set
        End Property

        Public ReadOnly Property colIsDirty() As Boolean
            Get
                Dim xAddressInfo As MUSTER.Info.AddressInfo

                For Each xAddressInfo In colAddresses.Values
                    If xAddressInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
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
                Return oAddressInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oAddressInfo.CreatedBy = Value
            End Set
        End Property

        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oAddressInfo.CreatedOn
            End Get
        End Property

        Public Property ModifiedBy() As String
            Get
                Return oAddressInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oAddressInfo.ModifiedBy = Value
            End Set
        End Property

        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oAddressInfo.ModifiedOn
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Sub Load(ByVal dr As DataRow)
            Try
                oAddressInfo = New MUSTER.Info.AddressInfo(dr)
                colAddresses.Add(oAddressInfo)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        Public Function Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal pEntityType As Integer, ByVal entity As Integer, Optional ByVal bolValidated As Boolean = False) As Boolean
            Dim oldID As Integer
            Try
                If Not bolValidated Then
                    If Not Me.ValidateData() Then
                        Return False
                    End If
                End If
                If Not (oAddressInfo.AddressId < 0 And oAddressInfo.Deleted) Then
                    oldID = oAddressInfo.AddressId
                    oAddressDB.Put(oAddressInfo, moduleID, staffID, returnVal, pEntityType, entity)
                    If Not returnVal = String.Empty Then
                        Exit Function
                    End If

                    If oldID <> oAddressInfo.AddressId Then
                        colAddresses.ChangeKey(oldID, oAddressInfo.AddressId)
                    End If
                    oAddressInfo.Archive()
                    oAddressInfo.IsDirty = False
                End If
                If oAddressInfo.Deleted Then
                    If Not bolValidated Then
                        ' check if other owners are present else load new instance
                        Dim strNext As String = Me.GetNext()
                        Dim strPrev As String = Me.GetPrevious()
                        If strNext = oAddressInfo.AddressId Then
                            If strPrev = oAddressInfo.AddressId Then
                                RaiseEvent evtAddressErr("Address " + oAddressInfo.AddressId.ToString + " deleted")
                                colAddresses.Remove(oAddressInfo)
                                oAddressInfo = Me.Retrieve(0)
                            Else
                                RaiseEvent evtAddressErr("Address " + oAddressInfo.AddressId.ToString + " deleted")
                                colAddresses.Remove(oAddressInfo)
                                oAddressInfo = Me.Retrieve(strPrev)
                            End If
                        Else
                            RaiseEvent evtAddressErr("Address " + oAddressInfo.AddressId.ToString + " deleted")
                            colAddresses.Remove(oAddressInfo)
                            oAddressInfo = Me.Retrieve(strNext)
                        End If
                    Else
                        RaiseEvent evtAddressErr("Address " + oAddressInfo.AddressId.ToString + " deleted")
                        colAddresses.Remove(oAddressInfo)
                    End If
                End If
                RaiseEvent evtAddressChanged(oAddressInfo.IsDirty)
                Return True
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Sub Save(ByRef oAddress As MUSTER.Info.AddressInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal pEntityType As Integer, ByVal entity As Integer)
            Dim oldID As Integer
            Try
                oldID = oAddress.AddressId
                oAddressDB.Put(oAddress, moduleID, staffID, returnVal, pEntityType, entity)
                If Not returnVal = String.Empty Then
                    Exit Sub
                End If

                If oldID <> oAddress.AddressId Then
                    colAddresses.ChangeKey(oldID, oAddress.AddressId)
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
        Public Function Retrieve(ByVal ID As Int64, Optional ByVal strDepth As String = "SELF", Optional ByVal showDeleted As Boolean = False, Optional ByVal bolLoading As Boolean = False) As MUSTER.Info.AddressInfo
            Try
                'If Not bolLoading Then
                '    Me.ValidateData()
                'End If
                If oAddressInfo.AddressId < 0 And _
                    Not oAddressInfo.IsDirty And _
                    ID = 0 Then
                    Exit Try
                End If
                Select Case UCase(strDepth).Trim
                    Case "SELF", "CHILD", "GRANDCHILD", "ALL"
                        ' retrieve address info only
                        oAddressInfo = colAddresses.Item(ID)
                        If Not (oAddressInfo Is Nothing) Then
                            'If data is unchanged and old, dump it and get new data
                            If oAddressInfo.IsDirty = False And oAddressInfo.IsAgedData = True Then
                                colAddresses.Remove(oAddressInfo)
                            Else
                                Exit Select
                            End If
                        End If
                        Add(ID, showDeleted)
                    Case Else
                        RaiseEvent evtAddressErr("Pass correct param for Address retrieve")
                End Select
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Return oAddressInfo
        End Function
        Function GetDataSet(ByVal strSQL As String) As DataSet
            Try
                Dim ds As DataSet
                ds = oAddressDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Validates Data according to DDD Specifications
        Public Function ValidateData() As Boolean
            Try
                Dim errStr As String = ""
                If oAddressInfo.AddressId <> 0 Then
                    If oAddressInfo.AddressLine1 = String.Empty Then
                        errStr += "AddressLine1 cannot be empty" + vbCrLf
                    End If
                    If oAddressInfo.City = String.Empty Then
                        errStr += "City cannot be empty" + vbCrLf
                    End If
                    If oAddressInfo.State = String.Empty Then
                        errStr += "State cannot be empty" + vbCrLf
                    End If
                    If oAddressInfo.Zip = String.Empty Then
                        errStr += "Zip cannot be empty" + vbCrLf
                    End If
                End If
                If errStr.Length > 0 Then
                    RaiseEvent evtAddressErr(errStr)
                    Return False
                End If
                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Function GetAddressAll(ByVal blnShowDeleted As Boolean) As MUSTER.Info.AddressCollection
            colAddresses.Clear()
            Try
                colAddresses = oAddressDB.GetAllInfo(blnShowDeleted)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Return colAddresses
        End Function
        Public Sub Add(ByVal ID As Int64, Optional ByVal showDeleted As Boolean = False)
            Try
                oAddressInfo = oAddressDB.DBGetByID(ID, showDeleted)
                If oAddressInfo.AddressId = 0 Then
                    oAddressInfo.AddressId = nID
                    nID -= 1
                End If
                colAddresses.Add(oAddressInfo)
            Catch ex As Exception
                Throw ex
            End Try

        End Sub
        Public Sub Add(ByRef oAddress As MUSTER.Info.AddressInfo)
            Try
                oAddressInfo = oAddress
                If oAddressInfo.AddressId = 0 Then
                    oAddressInfo.AddressId = nID
                    nID -= 1
                End If
                colAddresses.Add(oAddressInfo)
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Public Sub Remove(ByVal ID As Int64)
            Dim myIndex As Int16 = 1
            Dim oAddressInfoLocal As MUSTER.Info.AddressInfo
            Try
                oAddressInfoLocal = colAddresses.Item(ID)
                If Not (oAddressInfoLocal Is Nothing) Then
                    colAddresses.Remove(oAddressInfoLocal)
                    Exit Sub
                End If
            Catch ex As Exception
                Throw ex
            End Try
            'Throw New Exception("Address " & ID.ToString & " is not in the collection of addresses.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal entitytypeID As Integer, ByVal entity As Integer)
            Try
                Dim IDs As New Collection
                Dim delIDs As New Collection
                Dim index As Integer
                Dim oTempInfo As MUSTER.Info.AddressInfo
                For Each oTempInfo In colAddresses.Values
                    If oTempInfo.IsDirty Then
                        oAddressInfo = oTempInfo
                        If oAddressInfo.Deleted Then
                            If oAddressInfo.AddressId < 0 Then
                                delIDs.Add(oAddressInfo.AddressId)
                            Else
                                Me.Save(moduleID, staffID, returnVal, EntityType, entity, True)
                            End If
                        Else
                            If Me.ValidateData Then
                                If oAddressInfo.AddressId < 0 Then
                                    IDs.Add(oAddressInfo.AddressId)
                                End If
                                Me.Save(moduleID, staffID, returnVal, EntityType, entity, True)
                            Else : Exit For
                            End If
                        End If
                    End If
                Next
                If Not (delIDs Is Nothing) Then
                    For index = 1 To delIDs.Count
                        oTempInfo = colAddresses.Item(CType(delIDs.Item(index), String))
                        colAddresses.Remove(oTempInfo)
                    Next
                End If
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        oTempInfo = colAddresses.Item(colKey)
                        colAddresses.ChangeKey(colKey, oTempInfo.AddressId)
                    Next
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colAddresses.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.AddressId.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colAddresses.Item(nArr.GetValue(colIndex + direction)).AddressId.ToString
            Else
                Return colAddresses.Item(nArr.GetValue(colIndex)).AddressId.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Function Clear(Optional ByVal strDepth As String = "ALL")
            oAddressInfo = New MUSTER.Info.AddressInfo
        End Function
        Public Function Reset(Optional ByVal strDepth As String = "ALL")
            oAddressInfo.Reset()
        End Function
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the addresses in the collection
        Public Function AddressTable() As DataTable

            Dim oAddressInfoLocal As MUSTER.Info.AddressInfo
            Dim dr As DataRow
            Dim tbAddressTable As New DataTable

            Try
                tbAddressTable.Columns.Add("ADDRESS_ID")
                tbAddressTable.Columns.Add("ADDRESS_TYPE_ID")
                tbAddressTable.Columns.Add("ENTITY_TYPE")
                tbAddressTable.Columns.Add("ADDRESS_LINE_ONE")
                tbAddressTable.Columns.Add("ADDRESS_TWO")
                tbAddressTable.Columns.Add("CITY")
                tbAddressTable.Columns.Add("STATE")
                tbAddressTable.Columns.Add("ZIP")
                tbAddressTable.Columns.Add("FIPS_CODE")
                tbAddressTable.Columns.Add("START_DATE")
                tbAddressTable.Columns.Add("END_DATE")
                tbAddressTable.Columns.Add("DELETED")
                tbAddressTable.Columns.Add("CREATED_BY")
                tbAddressTable.Columns.Add("DATE_CREATED")
                tbAddressTable.Columns.Add("LAST_EDITED_BY")
                tbAddressTable.Columns.Add("DATE_LAST_EDITED")
                tbAddressTable.Columns.Add("PHYSICALTOWN")

                For Each oAddressInfoLocal In colAddresses.Values
                    dr = tbAddressTable.NewRow()
                    dr("ADDRESS_ID") = oAddressInfoLocal.AddressId
                    dr("ADDRESS_TYPE_ID") = oAddressInfoLocal.AddressTypeId
                    dr("ENTITY_TYPE") = oAddressInfoLocal.EntityType
                    dr("ADDRESS_LINE_ONE") = oAddressInfoLocal.AddressLine1
                    dr("ADDRESS_TWO") = oAddressInfoLocal.AddressLine2
                    dr("CITY") = oAddressInfoLocal.City
                    dr("STATE") = oAddressInfoLocal.State
                    dr("ZIP") = oAddressInfoLocal.Zip
                    dr("FIPS_CODE") = oAddressInfoLocal.FIPSCode
                    dr("START_DATE") = oAddressInfoLocal.StartDate
                    dr("END_DATE") = oAddressInfoLocal.EndDate
                    dr("DELETED") = oAddressInfoLocal.Deleted
                    dr("CREATED_BY") = oAddressInfoLocal.CreatedBy
                    dr("DATE_CREATED") = oAddressInfoLocal.CreatedOn
                    dr("LAST_EDITED_BY") = oAddressInfoLocal.ModifiedBy
                    dr("DATE_LAST_EDITED") = oAddressInfoLocal.ModifiedOn
                    dr("PHYSICALTOWN") = oAddressInfoLocal.PhsycalTown
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
        Private Sub AddressesChanged(ByVal strSrc As String) Handles colAddresses.AddressColChanged
            RaiseEvent evtAddressesChanged(Me.colIsDirty)
        End Sub
        Private Sub AddressChanged(ByVal bolValue As Boolean) Handles oAddressInfo.AddressInfoChanged
            RaiseEvent evtAddressChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
