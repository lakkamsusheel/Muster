'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pProvider
'   Provides the operations required to manipulate an Course object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         MR     5/22/05    Original class definition.
'
' Function          Description
' Retrieve(ID)      Returns an Info Object requested by the int arg ID
' Save()            Saves the Info Object
' GetAll()          Returns a collection with all the relevant information
' Add(ID)           Adds an Info Object identified by the int arg ID
'                   to the Providers Collection
' Add(Entity)       Adds the Entity passed as an argument
'                   to the Providers Collection
' Remove(ID)        Removes an Info Object identified by the int arg ID
'                   from the Providers Collection
' Remove(Entity)    Removes the Entity passed as an argument
'                   from the Providers Collection
' Flush()           Marshalls all modified/new Onwer Info objects in the
'                   Provider Collection to the repository
' EntityTable()     Returns a datatable containing all columns for the Entity
'                   objects in the Providers Collection
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
        Public Class pProvider
#Region "Public Events"
        Public Event ProviderErr(ByVal MsgStr As String)
        Public Event evtProviderErr(ByVal MsgStr As String)
#End Region
#Region "Private Member Variables"
        Private oProviderInfo As MUSTER.Info.ProviderInfo
        Private colProvider As MUSTER.Info.ProviderCollection
        Private oProviderDB As MUSTER.DataAccess.ProviderDB
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Course").ID
#End Region
#Region "Constructors"
        Public Sub New()
            oProviderInfo = New MUSTER.Info.ProviderInfo
            colProvider = New MUSTER.Info.ProviderCollection
            oProviderDB = New MUSTER.DataAccess.ProviderDB
        End Sub
#End Region
#Region "Exposed Attributes"

        Public Property ID() As Integer
            Get
                Return oProviderInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oProviderInfo.ID = Value
            End Set
        End Property
        Public Property Active() As Boolean
            Get
                Return oProviderInfo.Active
            End Get

            Set(ByVal value As Boolean)
                oProviderInfo.Active = value
            End Set
        End Property
        Public Property ProviderName() As String
            Get
                Return oProviderInfo.ProviderName
            End Get
            Set(ByVal Value As String)
                oProviderInfo.ProviderName = Value
            End Set
        End Property
        Public Property Abbrev() As String
            Get
                Return oProviderInfo.Abbrev
            End Get
            Set(ByVal Value As String)
                oProviderInfo.Abbrev = Value
            End Set
        End Property
        Public Property Department() As String
            Get
                Return oProviderInfo.Department
            End Get

            Set(ByVal value As String)
                oProviderInfo.Department = value
            End Set
        End Property
        Public Property Website() As String
            Get
                Return oProviderInfo.Website
            End Get

            Set(ByVal value As String)
                oProviderInfo.Website = value
            End Set
        End Property
        Public Property Deleted() As Integer
            Get
                Return oProviderInfo.Deleted
            End Get

            Set(ByVal value As Integer)
                oProviderInfo.Deleted = Integer.Parse(value)
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return oProviderInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oProviderInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oProviderInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oProviderInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oProviderInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oProviderInfo.ModifiedOn
            End Get
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oProviderInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oProviderInfo.IsDirty = value
            End Set
        End Property
        Public Property ProviderCollection() As MUSTER.Info.ProviderCollection
            Get
                Return colProvider
            End Get
            Set(ByVal Value As MUSTER.Info.ProviderCollection)
                colProvider = Value
            End Set
        End Property
        Public Property ProviderInfo() As MUSTER.Info.ProviderInfo
            Get
                Return oProviderInfo
            End Get
            Set(ByVal Value As MUSTER.Info.ProviderInfo)
                oProviderInfo = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.ProviderInfo
            Dim oProviderInfoLocal As MUSTER.Info.ProviderInfo
            Dim bolDataAged As Boolean = False

            Try
                For Each oProviderInfoLocal In colProvider.Values
                    If oProviderInfoLocal.ID = ID Then
                        If oProviderInfoLocal.IsDirty = False And oProviderInfoLocal.IsAgedData = True Then
                            bolDataAged = True
                        Else
                            oProviderInfo = oProviderInfoLocal
                            Return oProviderInfo
                        End If
                    End If
                Next
                If bolDataAged = True Then
                    colProvider.Remove(oProviderInfoLocal)
                End If
                oProviderInfo = oProviderDB.DBGetByID(ID)
                If oProviderInfo.ID = 0 Then
                    oProviderInfo.ID = nID
                    nID -= 1
                End If
                colProvider.Add(oProviderInfo)
                Return oProviderInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)

            Try
                Dim OldKey As String = oProviderInfo.ID.ToString
                oProviderDB.Put(oProviderInfo, moduleID, staffID, returnVal)
                If oProviderInfo.ID.ToString <> OldKey Then
                    colProvider.ChangeKey(OldKey, oProviderInfo.ID.ToString)
                End If
                oProviderInfo.Archive()
                oProviderInfo.IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        Public Sub Save(ByRef oProvInfo As MUSTER.Info.ProviderInfo, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)

            Try
                Dim OldKey As String = oProvInfo.ID.ToString
                oProviderDB.Put(oProvInfo, moduleID, staffID, returnVal)
                If oProvInfo.ID.ToString <> OldKey Then
                    colProvider.ChangeKey(OldKey, oProvInfo.ID.ToString)
                End If
                oProvInfo.Archive()
                oProvInfo.IsDirty = False
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        '' Validates Data according to DDD Specifications
        'Public Function ValidateData() As Boolean
        '    Try
        '        Dim errStr As String = ""
        '        Dim validateSuccess As Boolean = True

        '        If oProviderInfo.ID <> 0 Then
        '            If oProviderInfo.ProviderName <> String.Empty Then
        '                If oProviderInfo.Abbrev <> String.Empty Then
        '                    If oProviderInfo.Department <> String.Empty Then
        '                        validateSuccess = True
        '                    Else
        '                        errStr += "Department cannot be empty" + vbCrLf
        '                        validateSuccess = False
        '                    End If
        '                Else
        '                    errStr += "Abbreviation cannot be empty" + vbCrLf
        '                    validateSuccess = False
        '                End If
        '            Else
        '                errStr += "Provider cannot be empty" + vbCrLf
        '                validateSuccess = False
        '            End If
        '        End If
        '        If errStr.Length > 0 Or Not validateSuccess Then
        '            RaiseEvent evtProviderErr(errStr)
        '        End If
        '        Return validateSuccess
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function
#End Region
#Region "Collection Operations"
        Function GetAll() As MUSTER.Info.ProviderCollection
            Try
                colProvider.Clear()
                colProvider = oProviderDB.GetAllInfo()
                Return colProvider
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oProviderInfo = oProviderDB.DBGetByID(ID)
                If oProviderInfo.ID = 0 Then
                    oProviderInfo.ID = nID
                    nID -= 1
                End If
                colProvider.Add(oProviderInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Function Add(ByRef oProvider As MUSTER.Info.ProviderInfo) As Boolean
            Try
                'If ValidateData() Then
                oProviderInfo = oProvider
                colProvider.Add(oProviderInfo)
                'Return True
                'Else
                '    Return False
                'End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oProviderInfoLocal As MUSTER.Info.ProviderInfo

            Try
                For Each oProviderInfoLocal In colProvider.Values
                    If oProviderInfoLocal.ID = ID Then
                        colProvider.Remove(oProviderInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Provider " & ID.ToString & " is not in the collection of Providers.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oProvider As MUSTER.Info.ProviderInfo)
            Try
                colProvider.Remove(oProvider)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Provider " & oProvider.ID & " is not in the collection of Providers.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xProviderInfo As MUSTER.Info.ProviderInfo
            For Each xProviderInfo In colProvider.Values
                If xProviderInfo.IsDirty Then
                    oProviderInfo = xProviderInfo
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
        End Sub
#End Region

#Region "General Operations"
        Public Sub Clear()
            oProviderInfo = New MUSTER.Info.ProviderInfo
        End Sub
        Public Sub Reset()
            oProviderInfo.Reset()
        End Sub
#End Region
#Region "LookUp Operations"
        Public Function ListProviderNames(Optional ByVal showBlankPropertyName As Boolean = True) As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vCOM_PROVIDERNAME")
                If showBlankPropertyName Then
                    Dim dr As DataRow = dtReturn.NewRow
                    For Each dtCol As DataColumn In dtReturn.Columns
                        If dtCol.DataType.Name.IndexOf("String") > -1 Then
                            dr(dtCol) = " "
                        ElseIf dtCol.DataType.Name.IndexOf("Int") > -1 Then
                            dr(dtCol) = 0
                        End If
                    Next
                    dtReturn.Rows.InsertAt(dr, 0)
                End If

                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Private Function GetDataTable(ByVal DBViewName As String) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try
                strSQL = "SELECT * FROM " & DBViewName

                dsReturn = oProviderDB.DBGetDS(strSQL)
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
#End Region
#Region "Miscellaneous Operations"
#End Region
#End Region
    End Class
End Namespace
