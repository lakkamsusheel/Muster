'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.ZipCode
'   Provides the info and collection objects to the client for manipulating
'   an AddressInfo object.
'   Copyright (C) 2004 CIBER, Inc.
'  All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         EN      12/13/04    Original class definition.
'   1.1         EN      12/27/04    Changed from Put to AddZip in Save method.
'   1.2         EN      12/29/04    Added the properties in to collection while setting the property.
'   1.3         AN      01/03/05    Added Try catch and Exception Handling/Logging
'   1.4         MNR     01/25/05    Added GetNext(), GetPrevious() and GetNextPrev(..) functions
'   1.5         AB      02/22/05    Added DataAge check to the Retrieve function
'
' Function          Description
' Retrieve(FullKey)     Returns the Zipcode info requested by the Full Key
' GetAll()          Returns an  ZipCodeCollection with all ZipCode objects
' Add(ID)           Adds the ZipCode identified by arg ID to the 
'                           internal  ZipCodeCollection
' Remove(ID)        Removes the ZipCode identified by arg ID from the internal 
'                            ZipCodeCollection
' ZipCodeTable()     Returns a datatable containing all columns for the ZipCode 

' colIsDirty()       Returns a boolean indicating whether any of the LetterInfo
'                    objects in the ZipCodeCollection has been modified since the
'                    last time it was retrieved from/saved to the repository.
' Flush()            Marshalls all modified/added ZipCodeInfo objects in the 
'                        ZipCodeCollection to the repository.
' Save()             Marshalls the internal ZipCodeInfo object to the repository.
'CSZTable            returns datatable based on 

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pZipCode
#Region "Public Events"
        Public Event ZipCodeErr(ByVal MsgStr As String)
#End Region
#Region "Private Member Variables"
        Private ColZipCode As Muster.Info.ZipCodeCollection
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private oZipCodeInfo As Muster.Info.ZipCodeInfo
        Private oZipCodeDB As New Muster.DataAccess.ZipCodeDB
        Private strNewID As String
        Private nID As Int64 = -1
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            oZipCodeInfo = New Muster.Info.ZipCodeInfo
            ColZipCode = New Muster.Info.ZipCodeCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return oZipCodeInfo.ID
            End Get
            Set(ByVal value As String)
                oZipCodeInfo.ID = value
            End Set
        End Property
        Public Property County() As String
            Get
                Return oZipCodeInfo.County
            End Get
            Set(ByVal value As String)
                oZipCodeInfo.County = value
                ColZipCode(oZipCodeInfo.ID) = oZipCodeInfo

            End Set
        End Property
        Public Property Zip() As String
            Get
                Return oZipCodeInfo.Zip
            End Get
            Set(ByVal value As String)
                oZipCodeInfo.Zip = value
                ColZipCode(oZipCodeInfo.ID) = oZipCodeInfo
            End Set
        End Property
        Public Property City() As String
            Get
                Return oZipCodeInfo.City
            End Get
            Set(ByVal value As String)
                oZipCodeInfo.City = value
                ColZipCode(oZipCodeInfo.ID) = oZipCodeInfo
            End Set
        End Property
        Public Property state() As String
            Get
                Return oZipCodeInfo.state
            End Get
            Set(ByVal value As String)
                oZipCodeInfo.state = value
                ColZipCode(oZipCodeInfo.ID) = oZipCodeInfo
            End Set
        End Property
        Public Property FIPS() As String
            Get
                Return oZipCodeInfo.Fips
            End Get
            Set(ByVal value As String)
                oZipCodeInfo.Fips = value
                ColZipCode(oZipCodeInfo.ID) = oZipCodeInfo
            End Set
        End Property
        Public ReadOnly Property CreatedBy() As String
            Get
                Return oZipCodeInfo.CreatedBy
            End Get
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oZipCodeInfo.CreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedBy() As String
            Get
                Return oZipCodeInfo.ModifiedBy
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oZipCodeInfo.ModifiedOn
            End Get
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xZipCodeInfo As MUSTER.Info.ZipCodeInfo
                For Each xZipCodeInfo In ColZipCode.Values
                    If xZipCodeInfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        Public Function Retrieve(ByVal FullKey As String) As Muster.Info.ZipCodeInfo
            Try
                If ColZipCode.Contains(FullKey) Then
                    oZipCodeInfo = ColZipCode.Item(FullKey)
                    If oZipCodeInfo.IsAgedData = True And oZipCodeInfo.IsDirty = False Then
                        ColZipCode.Remove(oZipCodeInfo)
                    Else
                        Return oZipCodeInfo
                    End If
                End If

                Dim strArray() As String
                strArray = FullKey.Split("|")
                oZipCodeInfo = oZipCodeDB.DBGetByKey(strArray)
                Return oZipCodeInfo

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Function GetAll() As Muster.Info.ZipCodeCollection
            Try
                ColZipCode.Clear()
                ColZipCode = oZipCodeDB.GetAllInfo()
                Return ColZipCode
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an address to the collection as supplied by the caller
        Public Sub Add(ByRef oZipCode As Muster.Info.ZipCodeInfo)
            Try
                oZipCodeInfo = oZipCode
                'If oZipCodeInfo.ID = 0 Then
                oZipCodeInfo.ID = "0" & "|" & "0" & "|" & "O" & "|" & nID
                nID -= 1
                'End If

                ColZipCode.Add(oZipCodeInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        'Removes the address called for by ID from the collection
        Public Sub Remove(ByVal ID As String)
            Dim myIndex As Int16 = 1
            Try
                For Each oZipCodeInfo In ColZipCode.Values
                    If oZipCodeInfo.ID = Trim(ID) Then
                        ColZipCode.Remove(oZipCodeInfo)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

            Throw New Exception("Letter " & ID.ToString & " is not in the collection of Letter.")
        End Sub
        ''Public Function colIsDirty() As Boolean
        ''    Dim oTempInfo As Muster.Info.ZipCodeInfo
        ''    For Each oTempInfo In ColZipCode.Values
        ''        If oTempInfo.IsDirty Then
        ''            Return True
        ''        End If
        ''    Next
        ''    Return False
        ''End Function
        Public Sub Flush()
            Dim oTempInfo As Muster.Info.ZipCodeInfo
            For Each oTempInfo In ColZipCode.Values
                If oTempInfo.IsDirty Then
                    oZipCodeInfo = oTempInfo
                    Me.Save()
                End If
            Next
        End Sub
        Public Sub Save()
            Try
                oZipCodeDB.AddZip(oZipCodeInfo)
                oZipCodeInfo.IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Save(ByRef oZipCode As Muster.Info.ZipCodeInfo)
            Try
                strNewID = oZipCodeDB.AddZip(oZipCode)
                If strNewID <> "" Then
                    oZipCodeInfo.ID = strNewID
                End If
                oZipCode.IsDirty = False
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
            Dim strArr() As String = ColZipCode.GetKeys()
            'Dim nArr(strArr.GetUpperBound(0)) As Integer
            'Dim y As String
            'For Each y In strArr
            '    nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            'Next
            'nArr.Sort(nArr)
            colIndex = Array.BinarySearch(strArr, Me.ID.ToString)
            If colIndex + direction > -1 And _
                colIndex + direction <= strArr.GetUpperBound(0) Then
                Return ColZipCode.Item(strArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return ColZipCode.Item(strArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oZipCodeInfo = New MUSTER.Info.ZipCodeInfo
        End Sub
        Public Sub Reset()
            oZipCodeInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        'Returns a datatable of the addresses in the collection
        Public Function ZipCodeTable() As DataTable
            Dim oZipCodeInfoLocal As Muster.Info.ZipCodeInfo
            Dim dr As DataRow
            Dim tbZipCodeTable As New DataTable
            Try
                tbZipCodeTable.Columns.Add("ID")
                tbZipCodeTable.Columns.Add("ZIP")
                tbZipCodeTable.Columns.Add("STATE")
                tbZipCodeTable.Columns.Add("CITY")
                tbZipCodeTable.Columns.Add("COUNTY")
                tbZipCodeTable.Columns.Add("FIPS")
                tbZipCodeTable.Columns.Add("created_by")
                tbZipCodeTable.Columns.Add("date_created")
                tbZipCodeTable.Columns.Add("last_edited_by")
                tbZipCodeTable.Columns.Add("date_last_edited")

                For Each oZipCodeInfoLocal In ColZipCode.Values
                    dr = tbZipCodeTable.NewRow()
                    dr("ID") = oZipCodeInfoLocal.ID
                    dr("ZIP") = oZipCodeInfoLocal.Zip
                    dr("STATE") = oZipCodeInfoLocal.state
                    dr("CITY") = oZipCodeInfoLocal.City
                    dr("COUNTY") = oZipCodeInfoLocal.County
                    dr("FIPS") = oZipCodeInfoLocal.Fips
                    dr("CREATED_BY") = oZipCodeInfoLocal.CreatedBy
                    dr("date_created") = oZipCodeInfoLocal.CreatedOn
                    dr("last_edited_by") = oZipCodeInfoLocal.ModifiedBy
                    dr("date_last_edited") = oZipCodeInfoLocal.ModifiedOn
                    tbZipCodeTable.Rows.Add(dr)
                Next
                Return tbZipCodeTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function CSZTable(ByVal strZip As String) As DataTable
            Dim oAddressInfoLocal As Muster.Info.AddressInfo
            Dim tbAddressTable As New DataTable
            Try
                tbAddressTable = oZipCodeDB.DBGetDS("SELECT DISTINCT COUNTY,CITY,STATE,ZIP,FIPS,ZIPID from tblSYS_ZIPCODES where ZIP like '" & strZip.TrimEnd & "%' Order By Zip").Tables(0)
                Return tbAddressTable
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDataSet(ByVal strSQL As String) As DataSet
            Try
                Dim ds As DataSet
                ds = oZipCodeDB.DBGetDS(strSQL)
                Return ds
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function CheckZip(ByVal oZip As Muster.info.ZipCodeInfo) As Boolean
            Try
                Return oZipCodeDB.DBCheckZip(oZip)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
    End Class
End Namespace
