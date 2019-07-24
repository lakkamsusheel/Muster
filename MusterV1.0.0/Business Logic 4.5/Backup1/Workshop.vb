'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.Workshop
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
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
' NOTE: This file to be used as Workshop to build other objects.
'       Replace keyword "Workshop" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pWorkshop
#Region "Public Events"
        Public Event WorkshopErr(ByVal MsgStr As String)
        Public Event WorkshopChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oWorkshopInfo As Muster.Info.WorkshopInfo
        Private WithEvents colWorkshops As Muster.Info.WorkshopsCollection
        Private oWorkshopDB As New Muster.DataAccess.WorkshopDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
#End Region
#Region "Constructors"
        Public Sub New()
            oWorkshopInfo = New Muster.Info.WorkshopInfo
            colWorkshops = New Muster.Info.WorkshopsCollection
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named Workshop object.
        '
        '********************************************************
        Public Sub New(ByVal WorkshopName As String)
            oWorkshopInfo = New Muster.Info.WorkshopInfo
            colWorkshops = New Muster.Info.WorkshopsCollection
            Me.Retrieve(WorkshopName)
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oWorkshopInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oWorkshopInfo.ID = Integer.Parse(Value)
            End Set
        End Property
        Public Property Name() As String
            Get
                Return oWorkshopInfo.Name
            End Get
            Set(ByVal Value As String)
                oWorkshopInfo.Name = Value
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return oWorkshopInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oWorkshopInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oWorkshopInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oWorkshopInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xWorkshopinfo As Muster.Info.WorkshopInfo
                For Each xWorkshopinfo In colWorkshops.Values
                    If xWorkshopinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oWorkshopInfo.IsDirty = Value
            End Set
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As Muster.Info.WorkshopInfo
            Dim oWorkshopInfoLocal As Muster.Info.WorkshopInfo
            Try
                For Each oWorkshopInfoLocal In colWorkshops.Values
                    If oWorkshopInfoLocal.ID = ID Then
                        oWorkshopInfo = oWorkshopInfoLocal
                        Return oWorkshopInfo
                    End If
                Next
                oWorkshopInfo = oWorkshopDB.DBGetByID(ID)
                If oWorkshopInfo.ID = 0 Then
                    oWorkshopInfo.ID = nID
                    nID -= 1
                End If
                colWorkshops.Add(oWorkshopInfo)
                Return oWorkshopInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function Retrieve(ByVal WorkshopName As String) As Muster.Info.WorkshopInfo
            Try
                oWorkshopInfo = Nothing
                If colWorkshops.Contains(WorkshopName) Then
                    oWorkshopInfo = colWorkshops(WorkshopName)
                Else
                    If oWorkshopInfo Is Nothing Then
                        oWorkshopInfo = New Muster.Info.WorkshopInfo
                    End If
                    oWorkshopInfo = oWorkshopDB.DBGetByName(WorkshopName)
                    colWorkshops.Add(oWorkshopInfo)
                End If
                Return oWorkshopInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function
        'Saves the data in the current Info object
        Public Sub Save()
            Dim strModuleName As String = String.Empty
            Try
                If Me.ValidateData(strModuleName) Then
                    oWorkshopDB.Put(oWorkshopInfo)
                    oWorkshopInfo.Archive()
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = False
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
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"

                        ' if any validations failed
                        Exit Select
                End Select
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent WorkshopErr(errStr)
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
        Function GetAll() As MUSTER.Info.WorkshopsCollection
            Try
                colWorkshops.Clear()
                colWorkshops = oWorkshopDB.GetAllInfo
                Return colWorkshops
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oWorkshopInfo = oWorkshopDB.DBGetByID(ID)
                If oWorkshopInfo.ID = 0 Then
                    oWorkshopInfo.ID = nID
                    nID -= 1
                End If
                colWorkshops.Add(oWorkshopInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oWorkshop As MUSTER.Info.WorkshopInfo)
            Try
                oWorkshopInfo = oWorkshop
                colWorkshops.Add(oWorkshopInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oWorkshopInfoLocal As MUSTER.Info.WorkshopInfo

            Try
                For Each oWorkshopInfoLocal In colWorkshops.Values
                    If oWorkshopInfoLocal.ID = ID Then
                        colWorkshops.Remove(oWorkshopInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Workshop " & ID.ToString & " is not in the collection of Workshops.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oWorkshop As MUSTER.Info.WorkshopInfo)
            Try
                colWorkshops.Remove(oWorkshop)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Workshop " & oWorkshop.ID & " is not in the collection of Workshops.")
        End Sub
        Public Sub Flush()
            Dim xWorkshopInfo As MUSTER.Info.WorkshopInfo
            For Each xWorkshopInfo In colWorkshops.Values
                If xWorkshopInfo.IsDirty Then
                    oWorkshopInfo = xWorkshopInfo
                    '********************************************************
                    '
                    ' Note that if there are contained objects and the respective
                    '  contained object collections are dirty, then the contained
                    '  collections MUST BE FLUSHED before this object can be
                    '  saved.  Otherwise, there is a risk that an attempt will
                    '  be made to insert a new object to the repository without
                    '  corresponding contained information being present which 
                    '  may, in turn, cause a foreign key violation!
                    '
                    '********************************************************
                    Me.Save()
                End If
            Next
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colWorkshops.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colWorkshops.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colWorkshops.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oWorkshopInfo = New MUSTER.Info.WorkshopInfo
        End Sub
        Public Sub Reset()
            oWorkshopInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oWorkshopInfoLocal As New Muster.Info.WorkshopInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("Workshop ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oWorkshopInfoLocal In colWorkshops.Values
                    dr = tbEntityTable.NewRow()
                    dr("ReplaceID") = oWorkshopInfoLocal.ID
                    dr("Deleted") = oWorkshopInfoLocal.Deleted
                    dr("Created By") = oWorkshopInfoLocal.CreatedBy
                    dr("Date Created") = oWorkshopInfoLocal.CreatedOn
                    dr("Last Edited By") = oWorkshopInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oWorkshopInfoLocal.ModifiedOn
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
        Private Sub WorkshopInfoChanged(ByVal bolValue As Boolean) Handles oWorkshopInfo.WorkshopInfoChanged
            RaiseEvent WorkshopChanged(bolValue)
        End Sub
        Private Sub WorkshopColChanged(ByVal bolValue As Boolean) Handles colWorkshops.WorkshopColChanged
            RaiseEvent ColChanged(bolValue)
        End Sub
#End Region
    End Class
End Namespace
