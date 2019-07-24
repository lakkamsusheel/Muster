
Namespace MUSTER.BusinessLogic
    '  -------------------------------------------------------------------------------
    '  MUSTER.BusinessLogic.pPendFeeLine
    '  Provides the operations required to manipulate a Pending Fee Line object.
    '  
    '  Copyright (C) 2004, 2005 CIBER, Inc.
    '  All rights reserved.
    '  
    '  Release   Initials    Date        Description
    '  1.0         AN       06/28/2005    Original class definition
    '  
    '  Function          Description
    '  -------------------------------------------------------------------------------
    '  Attribute          Description
    '  -------------------------------------------------------------------------------
    Public Class pPendFeeLine

#Region "Public Events"
        Public Event PendFeeLineBLChanged As PendFeeLineBLChangedEventHandler
        Public Event PendFeeLineBLColChanged As PendFeeLineBLColChangedEventHandler
        Public Event PendFeeLineBLErr As PendFeeLineBLErrEventHandler
        Public Event PendFeeLineInfChanged As PendFeeLineInfoChanged

        Public Delegate Sub PendFeeLineBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub PendFeeLineBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub PendFeeLineBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub PendFeeLineInfoChanged()
#End Region
#Region "Private member variables"
        Private WithEvents oPendFeeLine As MUSTER.Info.PendFeeLineInfo
        Private WithEvents oPendFeeLineCol As New MUSTER.Info.PendFeeLineCollection
        Private oPendFeeLineDB As New MUSTER.DataAccess.PendFeeLineDB

        Private MusterExceptions As Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            oPendFeeLineDB = New MUSTER.dataaccess.PendFeeLineDB
            oPendFeeLineCol = New MUSTER.Info.PendFeeLineCollection
        End Sub
        Public Sub New(ByVal strName As String)
            oPendFeeLineDB = New MUSTER.dataaccess.PendFeeLineDB
            oPendFeeLineCol = New MUSTER.Info.PendFeeLineCollection
        End Sub
        Public Sub New(ByVal PendFeeLineID As Integer)
            oPendFeeLineDB = New MUSTER.dataaccess.PendFeeLineDB
            oPendFeeLineCol = New MUSTER.Info.PendFeeLineCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return oPendFeeLine.ID
            End Get
            Set(ByVal Value As Integer)
                oPendFeeLine.ID = Value
            End Set
        End Property

        Public Property InvoiceAdviceId() As Integer
            Get
                Return oPendFeeLine.InvoiceAdviceId
            End Get
            Set(ByVal Value As Integer)
                oPendFeeLine.InvoiceAdviceId = Value
            End Set
        End Property

        Public Property ItemSequenceNumber() As Integer
            Get
                Return oPendFeeLine.ItemSequenceNumber
            End Get
            Set(ByVal Value As Integer)
                Value = oPendFeeLine.ItemSequenceNumber
            End Set
        End Property

        Public Property InvoiceNumber() As String
            Get
                Return oPendFeeLine.InvoiceNumber
            End Get
            Set(ByVal Value As String)
                oPendFeeLine.InvoiceNumber = Value
            End Set
        End Property

        Public Property FacilityId() As Integer
            Get
                Return oPendFeeLine.FacilityId
            End Get
            Set(ByVal Value As Integer)
                oPendFeeLine.FacilityId = Value
            End Set
        End Property

        Public Property OwnerId() As Integer
            Get
                Return oPendFeeLine.OwnerId
            End Get
            Set(ByVal Value As Integer)
                oPendFeeLine.OwnerId = Value
            End Set
        End Property

        Public Property FiscalYear() As Integer
            Get
                Return oPendFeeLine.FiscalYear
            End Get
            Set(ByVal Value As Integer)
                oPendFeeLine.FiscalYear = Value
            End Set
        End Property

        Public Property InvoiceDate() As DateTime
            Get
                Return oPendFeeLine.InvoiceDate
            End Get
            Set(ByVal Value As DateTime)
                oPendFeeLine.InvoiceDate = Value
            End Set
        End Property

        Public Property Quantity() As Integer
            Get
                Return oPendFeeLine.Quantity
            End Get
            Set(ByVal Value As Integer)
                oPendFeeLine.Quantity = Value
            End Set
        End Property

        Public Property UnitPrice() As Decimal
            Get
                Return oPendFeeLine.UnitPrice
            End Get
            Set(ByVal Value As Decimal)
                oPendFeeLine.UnitPrice = Value
            End Set
        End Property

        Public Property InvoiceLineAmount() As Integer
            Get
                Return oPendFeeLine.InvoiceLineAmount
            End Get
            Set(ByVal Value As Integer)
                oPendFeeLine.InvoiceLineAmount = Value
            End Set
        End Property

        Public Property FeeType() As String
            Get
                Return oPendFeeLine.FeeType
            End Get
            Set(ByVal Value As String)
                oPendFeeLine.FeeType = Value
            End Set
        End Property

        Public Property InvoiceType() As Integer
            Get
                Return oPendFeeLine.InvoiceType
            End Get
            Set(ByVal Value As Integer)
                oPendFeeLine.InvoiceType = Value
            End Set
        End Property

        Public Property DueDate() As DateTime
            Get
                Return oPendFeeLine.DueDate
            End Get
            Set(ByVal Value As DateTime)
                oPendFeeLine.DueDate = Value
            End Set
        End Property

        Public Property Description() As String
            Get
                Return oPendFeeLine.Description
            End Get
            Set(ByVal Value As String)
                oPendFeeLine.Description = Value
            End Set
        End Property

        Public Property colIsDirty() As Boolean
            Get
            End Get
            Set(ByVal Value As Boolean)
            End Set
        End Property
        ' The ID of the user that created the row
        Public ReadOnly Property CreatedBy() As String
            Get

            End Get
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oPendFeeLine.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oPendFeeLine.Deleted = Value
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oPendFeeLine.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oPendFeeLine.IsDirty = Value
            End Set
        End Property

        Public ReadOnly Property ModifiedBy() As String
            Get
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
            End Get
        End Property
#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal PendFeeLineID As Int64) As MUSTER.Info.PendFeeLineInfo
            Dim oPendFeeLineInfoLocal As MUSTER.Info.PendFeeLineInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oPendFeeLineInfoLocal In oPendFeeLineCol.Values
                    If oPendFeeLineInfoLocal.ID = ID Then
                        If oPendFeeLineInfoLocal.IsAgedData = True And oPendFeeLineInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oPendFeeLine = oPendFeeLineInfoLocal
                            Return oPendFeeLine
                        End If
                    End If
                Next
                If bolDataAged Then
                    oPendFeeLineCol.Remove(oPendFeeLineInfoLocal)
                End If
                oPendFeeLine = oPendFeeLineDB.DBGetByID(PendFeeLineID)
                oPendFeeLineCol.Add(oPendFeeLine)
                Return oPendFeeLine
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Obtains and returns an entity as called for by ID
        Public Function GetByFiscalYear(ByVal PendFeeLineYear As Int32) As MUSTER.Info.PendFeeLineInfo
            Dim oPendFeeLineInfoLocal As MUSTER.Info.PendFeeLineInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oPendFeeLineInfoLocal In oPendFeeLineCol.Values
                    If oPendFeeLineInfoLocal.FiscalYear = PendFeeLineYear Then
                        If oPendFeeLineInfoLocal.IsAgedData = True And oPendFeeLineInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oPendFeeLine = oPendFeeLineInfoLocal
                            Return oPendFeeLine
                        End If
                    End If
                Next
                If bolDataAged Then
                    oPendFeeLineCol.Remove(oPendFeeLineInfoLocal)
                End If
                oPendFeeLine = oPendFeeLineDB.DBGetByFiscalYear(PendFeeLineYear)
                oPendFeeLineCol.Add(oPendFeeLine)
                Return oPendFeeLine
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True

            Try
                
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        ' Saves the data in the current Info object
        Public Sub Save(Optional ByVal strModuleName As String = "")
            Try
                If ValidateData() Then
                    oPendFeeLineDB.put(oPendFeeLine)
                    oPendFeeLine.IsDirty = False
                    oPendFeeLine.Archive()
                    RaiseEvent PendFeeLineBLChanged(oPendFeeLine.IsDirty)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Collection Operations"
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                ''oPendFeeLine = oPendFeeLineDB.DBGetByID(ID)
                oPendFeeLine.ID = ID
                If oPendFeeLine.ID = 0 Then
                    'oPendFeeLine.ID = nID
                    'nID -= 1
                End If
                oPendFeeLineCol.Add(oPendFeeLine)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef PendFeeLine As MUSTER.Info.PendFeeLineInfo)
            Try
                oPendFeeLine = PendFeeLine
                'oPendFeeLine.UserID = onUserID
                oPendFeeLineCol.Add(oPendFeeLine)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oPendFeeLineLocal As MUSTER.Info.PendFeeLineInfo

            Try
                For Each oPendFeeLineLocal In oPendFeeLineCol.Values
                    If oPendFeeLineLocal.ID = ID Then
                        oPendFeeLineCol.Remove(oPendFeeLineLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Pipe " & ID.ToString & " is not in the collection of Pipes.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLustEvent As MUSTER.Info.PendFeeLineInfo)
            Try
                oPendFeeLineCol.Remove(oLustEvent)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LustEvent " & oLustEvent.ID & " is not in the collection of PendFeeLine.")
        End Sub
        Public Sub Flush()
            Dim xPendFeeLineInfo As MUSTER.Info.PendFeeLineInfo
            For Each xPendFeeLineInfo In oPendFeeLineCol.Values
                If xPendFeeLineInfo.IsDirty Then
                    oPendFeeLine = xPendFeeLineInfo
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
            '    Dim strArr() As String = oPendFeeLineCol.GetKeys()
            '    Dim nArr(strArr.GetUpperBound(0)) As Integer
            '    Dim y As String
            '    For Each y In strArr
            '        nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            '    Next
            '    nArr.Sort(nArr)
            '    colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            '    If colIndex + direction > -1 And _
            '        colIndex + direction <= nArr.GetUpperBound(0) Then
            '        Return oPendFeeLineCol.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            '    Else
            '        Return oPendFeeLineCol.Item(nArr.GetValue(colIndex)).ID.ToString
            '    End If
        End Function
#End Region
#Region "General Operations"

        Public Function Reset() As Object
            oPendFeeLine.Reset()
            oPendFeeLineCol.Clear()
        End Function
#End Region
#Region "Miscellaneous Operations"

#End Region
#End Region
#Region "External Event Handlers"

#End Region
    End Class
End Namespace

