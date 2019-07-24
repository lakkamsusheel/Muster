'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.CompanyLiensee
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MKK/RAF    05/26/2005  Original class definition
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
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pCompanyLicensee
#Region "Public Events"
        Public Event CompanyLicenseeErr(ByVal MsgStr As String)
        Public Event CompanyLicenseeChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oCompanyLicenseeInfo As Muster.Info.CompanyLicenseeInfo
        Private WithEvents colCompanyLicensee As Muster.Info.CompanyLicenseeCollection
        Private oCompanyLicenseeDB As New Muster.DataAccess.CompanyLicenseeDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oCompanyLicenseeInfo = New Muster.Info.CompanyLicenseeInfo
            colCompanyLicensee = New Muster.Info.CompanyLicenseeCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named CompanyLicensee object.
        '
        '********************************************************
        Public Sub New(ByVal CompanyLicenseeName As String)
            oCompanyLicenseeInfo = New Muster.Info.CompanyLicenseeInfo
            colCompanyLicensee = New Muster.Info.CompanyLicenseeCollection
            Me.Retrieve(CompanyLicenseeName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oCompanyLicenseeInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oCompanyLicenseeInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oCompanyLicenseeInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oCompanyLicenseeInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oCompanyLicenseeInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oCompanyLicenseeInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xCompanyLicenseeinfo As MUSTER.Info.CompanyLicenseeInfo
                For Each xCompanyLicenseeinfo In colCompanyLicensee.Values
                    If xCompanyLicenseeinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oCompanyLicenseeInfo.IsDirty = Value
            End Set
        End Property
        Public Property CompanyID() As Integer
            Get
                Return oCompanyLicenseeInfo.CompanyID
            End Get
            Set(ByVal Value As Integer)
                oCompanyLicenseeInfo.CompanyID = Value
            End Set
        End Property
        Public Property LicenseeID() As Integer
            Get
                Return oCompanyLicenseeInfo.LicenseeID
            End Get
            Set(ByVal Value As Integer)
                oCompanyLicenseeInfo.LicenseeID = Value
            End Set
        End Property
        Public Property ComLicAddressID() As Integer
            Get
                Return oCompanyLicenseeInfo.ComLicAddressID
            End Get
            Set(ByVal Value As Integer)
                oCompanyLicenseeInfo.ComLicAddressID = Value
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return oCompanyLicenseeInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oCompanyLicenseeInfo.CreatedBy = Value
            End Set
        End Property
        Public Property CreatedOn() As DateTime
            Get
                Return oCompanyLicenseeInfo.CreatedOn
            End Get
            Set(ByVal Value As DateTime)
                oCompanyLicenseeInfo.CreatedOn = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oCompanyLicenseeInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oCompanyLicenseeInfo.ModifiedBy = Value
            End Set
        End Property
        Public Property ModifiedOn() As DateTime
            Get
                Return oCompanyLicenseeInfo.ModifiedOn
            End Get
            Set(ByVal Value As DateTime)
                oCompanyLicenseeInfo.ModifiedOn = Value
            End Set
        End Property
        Public Property ComLicCollection() As MUSTER.Info.CompanyLicenseeCollection
            Get
                Return colCompanyLicensee
            End Get
            Set(ByVal Value As MUSTER.Info.CompanyLicenseeCollection)
                colCompanyLicensee = Value
            End Set
        End Property
        Public Property ComLicInfo() As MUSTER.Info.CompanyLicenseeInfo
            Get
                Return oCompanyLicenseeInfo
            End Get
            Set(ByVal Value As MUSTER.Info.CompanyLicenseeInfo)
                oCompanyLicenseeInfo = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As Muster.Info.CompanyLicenseeInfo
            Dim oCompanyLicenseeInfoLocal As Muster.Info.CompanyLicenseeInfo
            Try
                For Each oCompanyLicenseeInfoLocal In colCompanyLicensee.Values
                    If oCompanyLicenseeInfoLocal.ID = ID Then
                        oCompanyLicenseeInfo = oCompanyLicenseeInfoLocal
                        Return oCompanyLicenseeInfo
                    End If
                Next
                oCompanyLicenseeInfo = oCompanyLicenseeDB.DBGetByAssociationID(ID)
                If oCompanyLicenseeInfo.ID = 0 Then
                    oCompanyLicenseeInfo.ID = nID
                    nID -= 1
                End If
                colCompanyLicensee.Add(oCompanyLicenseeInfo)
                Return oCompanyLicenseeInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal CompanyLicenseeName As String) As Muster.Info.CompanyLicenseeInfo
            Try
                oCompanyLicenseeInfo = Nothing
                If colCompanyLicensee.Contains(CompanyLicenseeName) Then
                    oCompanyLicenseeInfo = colCompanyLicensee(CompanyLicenseeName)
                Else
                    If oCompanyLicenseeInfo Is Nothing Then
                        oCompanyLicenseeInfo = New Muster.Info.CompanyLicenseeInfo
                    End If
                    oCompanyLicenseeInfo = oCompanyLicenseeDB.DBGetByAssociationID(CompanyLicenseeName)
                    colCompanyLicensee.Add(oCompanyLicenseeInfo)
                End If
                Return oCompanyLicenseeInfo
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
                    End If
                End If
                If Not ((oCompanyLicenseeInfo.ID < 0 And oCompanyLicenseeInfo.ID > -100) And oCompanyLicenseeInfo.Deleted) Then
                    OldKey = oCompanyLicenseeInfo.ID.ToString
                    oCompanyLicenseeDB.Put(oCompanyLicenseeInfo, moduleID, staffID, returnVal)
                    If Not bolValidated Then
                        If oCompanyLicenseeInfo.ID.ToString <> OldKey Then
                            colCompanyLicensee.ChangeKey(OldKey, oCompanyLicenseeInfo.ID.ToString)
                        End If
                    End If
                    oCompanyLicenseeInfo.Archive()
                    oCompanyLicenseeInfo.IsDirty = False
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True
            '********************************************************
            '
            ' Sample validation code below.  Modify as necessary to
            '  handle rules for the object.  Note that errStr should
            '  be built with all failed validation reasons as it is
            '  raised to the consumer for display to the user.  The
            '  same goes for the boolean validateSuccess.  Since
            '  validateSuccess ASSUMES failure, it must be set to
            '  TRUE if all validations are passed successfully.
            '********************************************************
            Try
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent CompanyLicenseeErr(errStr)
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
        Function GetAll(ByVal compID As Integer, Optional ByVal showDeleted As Boolean = False, Optional ByVal licenseeID As Integer = 0) As MUSTER.Info.CompanyLicenseeCollection
            Try
                colCompanyLicensee.Clear()
                colCompanyLicensee = oCompanyLicenseeDB.GetAssociations(compID, showDeleted, licenseeID)
                Return colCompanyLicensee
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oCompanyLicenseeInfo = oCompanyLicenseeDB.DBGetByAssociationID(ID)
                If oCompanyLicenseeInfo.ID = 0 Then
                    oCompanyLicenseeInfo.ID = nID
                    nID -= 1
                End If
                colCompanyLicensee.Add(oCompanyLicenseeInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Function Add(ByRef oCompanyLicensee As MUSTER.Info.CompanyLicenseeInfo) As Boolean
            Try
                oCompanyLicenseeInfo = oCompanyLicensee
                If ValidateData() Then
                    If oCompanyLicenseeInfo.ID = 0 Then
                        oCompanyLicenseeInfo.ID = nID
                        nID -= 1
                    End If
                    colCompanyLicensee.Add(oCompanyLicenseeInfo)
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
            Dim oCompanyLicenseeInfoLocal As MUSTER.Info.CompanyLicenseeInfo
            Try
                For Each oCompanyLicenseeInfoLocal In colCompanyLicensee.Values
                    If oCompanyLicenseeInfoLocal.ID = ID Then
                        colCompanyLicensee.Remove(oCompanyLicenseeInfoLocal)
                        Exit Sub
                    End If
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Association  ID " & oCompanyLicenseeInfo.ID & " is not in the collection of  Company Licensee Association.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oCompanyLicensee As MUSTER.Info.CompanyLicenseeInfo)
            Try
                colCompanyLicensee.Remove(oCompanyLicensee)
                Exit Sub
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Association  ID " & oCompanyLicenseeInfo.ID & " is not in the collection of  Company Licensee Association.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal id As Integer = 0)
            Dim index As Integer
            Dim IDs As New Collection
            Dim xCompanyLicenseeInfo As MUSTER.Info.CompanyLicenseeInfo
            Try
                For Each xCompanyLicenseeInfo In colCompanyLicensee.Values
                    If xCompanyLicenseeInfo.IsDirty Then
                        oCompanyLicenseeInfo = xCompanyLicenseeInfo
                        If id > 0 Or id < -100 Then
                            oCompanyLicenseeInfo.CompanyID = id
                        End If
                        IDs.Add(oCompanyLicenseeInfo.ID)
                        Me.Save(moduleID, staffID, returnVal, True)
                    End If
                Next
                If Not (IDs Is Nothing) Then
                    For index = 1 To IDs.Count
                        Dim colKey As String = CType(IDs.Item(index), String)
                        xCompanyLicenseeInfo = colCompanyLicensee.Item(colKey)
                        colCompanyLicensee.ChangeKey(colKey, xCompanyLicenseeInfo.ID)
                    Next
                End If
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "General Operations"
        Public Sub Clear()
            oCompanyLicenseeInfo = New MUSTER.Info.CompanyLicenseeInfo
        End Sub
        Public Sub Reset()
            oCompanyLicenseeInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oCompanyLicenseeInfoLocal As New MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("CompanyLicensee ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oCompanyLicenseeInfoLocal In colCompanyLicensee.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oCompanyLicenseeInfoLocal.ID
                    dr("Deleted") = oCompanyLicenseeInfoLocal.Deleted
                    dr("Created By") = oCompanyLicenseeInfoLocal.CreatedBy
                    dr("Date Created") = oCompanyLicenseeInfoLocal.CreatedOn
                    dr("Last Edited By") = oCompanyLicenseeInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oCompanyLicenseeInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub CompanyLicenseeInfoChanged(ByVal bolValue As Boolean) Handles oCompanyLicenseeInfo.CompanyLicenseeInfoChanged
            RaiseEvent CompanyLicenseeChanged(bolValue)
        End Sub
        Private Sub TamplateColChanged(ByVal bolValue As Boolean) Handles colCompanyLicensee.CompanyLicenseeColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
