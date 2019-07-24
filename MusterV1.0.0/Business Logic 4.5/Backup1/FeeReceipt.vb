


Namespace MUSTER.BusinessLogic
    '  -------------------------------------------------------------------------------
    '  MUSTER.BusinessLogic.pFeeReceipt
    '  Provides the operations required to manipulate a Fee Receipt object.
    '  
    '  Copyright (C) 2004, 2005 CIBER, Inc.
    '  All rights reserved.
    '  
    '  Release   Initials    Date        Description
    '  1.0         AB       09/26/2005    Original class definition
    '  
    '  Function          Description
    '  -------------------------------------------------------------------------------
    '  Attribute          Description
    '  -------------------------------------------------------------------------------
    Public Class pFeeReceipt

#Region "Public Events"
        Public Event FeeReceiptBLChanged As FeeReceiptBLChangedEventHandler
        Public Event FeeReceiptBLColChanged As FeeReceiptBLColChangedEventHandler
        Public Event FeeReceiptBLErr As FeeReceiptBLErrEventHandler
        Public Event FeeReceiptInfChanged As FeeReceiptInfoChanged

        Public Delegate Sub FeeReceiptBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeReceiptBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeReceiptBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FeeReceiptInfoChanged()
#End Region
#Region "Private member variables"
        Private WithEvents oFeeReceipt As MUSTER.Info.FeeReceiptInfo
        Private WithEvents oFeeReceiptCol As New MUSTER.Info.FeeReceiptCollection
        Private oFeeReceiptDB As New MUSTER.DataAccess.FeeReceiptDB

        Private MusterExceptions As Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            oFeeReceiptDB = New MUSTER.dataaccess.FeeReceiptDB
            oFeeReceiptCol = New MUSTER.Info.FeeReceiptCollection
        End Sub
        Public Sub New(ByVal strName As String)
            oFeeReceiptDB = New MUSTER.dataaccess.FeeReceiptDB
            oFeeReceiptCol = New MUSTER.Info.FeeReceiptCollection
        End Sub
        Public Sub New(ByVal FeeReceiptID As Integer)
            oFeeReceiptDB = New MUSTER.dataaccess.FeeReceiptDB
            oFeeReceiptCol = New MUSTER.Info.FeeReceiptCollection
        End Sub
#End Region
#Region "Exposed Attributes"

        Public ReadOnly Property ID() As Int64
            Get
                Return oFeeReceipt.ID
            End Get
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oFeeReceipt.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFeeReceipt.IsDirty = Value
            End Set
        End Property

        Public Property Deleted() As Boolean
            Get
                Return oFeeReceipt.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFeeReceipt.Deleted = Value
            End Set
        End Property
        Public Property FiscalYear() As String
            Get
                Return oFeeReceipt.FiscalYear
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.FiscalYear = Value
            End Set
        End Property
        Public Property ReturnType() As String
            Get
                Return oFeeReceipt.ReturnType
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.ReturnType = Value
            End Set
        End Property
        Public Property CheckTransID() As String
            Get
                Return oFeeReceipt.CheckTransID
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.CheckTransID = Value
            End Set
        End Property
        Public Property OwnerID() As Int64
            Get
                Return oFeeReceipt.OwnerID
            End Get
            Set(ByVal Value As Int64)
                oFeeReceipt.OwnerID = Value
            End Set
        End Property
        Public Property FacilityID() As Int64
            Get
                Return oFeeReceipt.FacilityID
            End Get
            Set(ByVal Value As Int64)
                oFeeReceipt.FacilityID = Value
            End Set
        End Property
        Public Property InvoiceNumber() As String
            Get
                Return oFeeReceipt.InvoiceNumber
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.InvoiceNumber = Value
            End Set
        End Property
        Public Property CheckNumber() As String
            Get
                Return oFeeReceipt.CheckNumber
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.CheckNumber = Value
            End Set
        End Property
        Public Property MisapplyFlag() As String
            Get
                Return oFeeReceipt.MisapplyFlag
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.MisapplyFlag = Value
            End Set
        End Property
        Public Property OverpaymentReason() As String
            Get
                Return oFeeReceipt.OverpaymentReason
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.OverpaymentReason = Value
            End Set
        End Property
        Public Property MisapplyReason() As String
            Get
                Return oFeeReceipt.MisapplyReason
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.MisapplyReason = Value
            End Set
        End Property
        Public Property IssuingCompany() As String
            Get
                Return oFeeReceipt.IssuingCompany
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.IssuingCompany = Value
            End Set
        End Property
        Public Property SequenceNumber() As Int16
            Get
                Return oFeeReceipt.SequenceNumber
            End Get
            Set(ByVal Value As Int16)
                oFeeReceipt.SequenceNumber = Value
            End Set
        End Property
        Public Property AmountReceived() As Single
            Get
                Return oFeeReceipt.AmountReceived
            End Get
            Set(ByVal Value As Single)
                oFeeReceipt.AmountReceived = Value
            End Set
        End Property
        Public Property ReceiptDate() As Date
            Get
                Return oFeeReceipt.ReceiptDate
            End Get
            Set(ByVal Value As Date)
                oFeeReceipt.ReceiptDate = Value
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return oFeeReceipt.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFeeReceipt.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFeeReceipt.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFeeReceipt.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFeeReceipt.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal FeeReceiptID As Int64) As MUSTER.Info.FeeReceiptInfo
            Dim oFeeReceiptInfoLocal As MUSTER.Info.FeeReceiptInfo
            Dim bolDataAged As Boolean = False
            Try
                'For Each oFeeReceiptInfoLocal In oFeeReceiptCol.Values
                '    If oFeeReceiptInfoLocal.ID = ID Then
                '        If oFeeReceiptInfoLocal.IsAgedData = True And oFeeReceiptInfoLocal.IsDirty = False Then
                '            bolDataAged = True
                '            Exit For
                '        Else
                '            oFeeReceipt = oFeeReceiptInfoLocal
                '            Return oFeeReceipt
                '        End If
                '    End If
                'Next
                'If bolDataAged Then
                '    oFeeReceiptCol.Remove(oFeeReceiptInfoLocal)
                'End If
                oFeeReceipt = oFeeReceiptDB.DBGetByID(FeeReceiptID)
                'oFeeReceiptCol.Add(oFeeReceipt)
                Return oFeeReceipt
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Obtains and returns an entity as called for by ID
        Public Function GetByFiscalYear(ByVal FeeReceiptYear As Int32) As MUSTER.Info.FeeReceiptInfo
            Dim oFeeReceiptInfoLocal As MUSTER.Info.FeeReceiptInfo
            Dim bolDataAged As Boolean = False
            Try
                'For Each oFeeReceiptInfoLocal In oFeeReceiptCol.Values
                '    If oFeeReceiptInfoLocal.FiscalYear = FeeReceiptYear Then
                '        If oFeeReceiptInfoLocal.IsAgedData = True And oFeeReceiptInfoLocal.IsDirty = False Then
                '            bolDataAged = True
                '            Exit For
                '        Else
                '            oFeeReceipt = oFeeReceiptInfoLocal
                '            Return oFeeReceipt
                '        End If
                '    End If
                'Next
                'If bolDataAged Then
                '    oFeeReceiptCol.Remove(oFeeReceiptInfoLocal)
                'End If
                oFeeReceipt = oFeeReceiptDB.DBGetByFiscalYear(FeeReceiptYear)
                'oFeeReceiptCol.Add(oFeeReceipt)
                Return oFeeReceipt
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
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "")
            Try
                If ValidateData() Then
                    oFeeReceiptDB.put(oFeeReceipt, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFeeReceipt.IsDirty = False
                    oFeeReceipt.Archive()
                    RaiseEvent FeeReceiptBLChanged(oFeeReceipt.IsDirty)
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
                ''oFeeReceipt = oFeeReceiptDB.DBGetByID(ID)
                oFeeReceipt.ID = ID
                If oFeeReceipt.ID = 0 Then
                    'oFeeReceipt.ID = nID
                    'nID -= 1
                End If
                oFeeReceiptCol.Add(oFeeReceipt)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef FeeReceipt As MUSTER.Info.FeeReceiptInfo)
            Try
                oFeeReceipt = FeeReceipt
                'oFeeReceipt.UserID = onUserID
                oFeeReceiptCol.Add(oFeeReceipt)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oFeeReceiptLocal As MUSTER.Info.FeeReceiptInfo

            Try
                For Each oFeeReceiptLocal In oFeeReceiptCol.Values
                    If oFeeReceiptLocal.ID = ID Then
                        oFeeReceiptCol.Remove(oFeeReceiptLocal)
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
        Public Sub Remove(ByVal oLustEvent As MUSTER.Info.FeeReceiptInfo)
            Try
                oFeeReceiptCol.Remove(oLustEvent)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LustEvent " & oLustEvent.ID & " is not in the collection of FeeReceipt.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xFeeReceiptInfo As MUSTER.Info.FeeReceiptInfo
            For Each xFeeReceiptInfo In oFeeReceiptCol.Values
                If xFeeReceiptInfo.IsDirty Then
                    oFeeReceipt = xFeeReceiptInfo
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
                    If oFeeReceipt.ID <= 0 Then
                        oFeeReceipt.CreatedBy = UserID
                    Else
                        oFeeReceipt.ModifiedBy = UserID
                    End If
                    Me.Save(moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
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
            '    Dim strArr() As String = oFeeReceiptCol.GetKeys()
            '    Dim nArr(strArr.GetUpperBound(0)) As Integer
            '    Dim y As String
            '    For Each y In strArr
            '        nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            '    Next
            '    nArr.Sort(nArr)
            '    colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            '    If colIndex + direction > -1 And _
            '        colIndex + direction <= nArr.GetUpperBound(0) Then
            '        Return oFeeReceiptCol.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            '    Else
            '        Return oFeeReceiptCol.Item(nArr.GetValue(colIndex)).ID.ToString
            '    End If
        End Function
#End Region
#Region "General Operations"

        Public Function Reset() As Object
            oFeeReceipt.Reset()
            oFeeReceiptCol.Clear()
        End Function
#End Region
#Region "Miscellaneous Operations"

#End Region

#Region " Populate Routines "
        'vFees_OwnerCheckTotals
        Public Function GetCheckTotal(ByVal OwnerID As Int64, ByVal CheckNumber As String) As Single
            Dim dsRemSys As New DataSet
            Dim dtReturn As Single
            Dim strSQL As String

            Try

                strSQL = "select Amount from vFees_OwnerCheckTotals where Owner_ID =  " & OwnerID & " and Check_Number = '" & CheckNumber & "'"

                dsRemSys = oFeeReceiptDB.DBGetDS(strSQL)
                dtReturn = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0).Rows(0)("Amount")
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetOwnerSummaryReceiptGrid_ByOwnerID(ByVal OwnerID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim dsRel1 As DataRelation
            Dim dsRel2 As DataRelation
            Dim strSQL As String


            strSQL = "select distinct * from vFees_OwnerReceiptSummaryGrid where Owner_ID = " & OwnerID & ";"
            strSQL = strSQL & " "
            strSQL = strSQL & "select distinct * from vFees_OwnerReceiptSummaryLineItemGrid where Owner_ID = " & OwnerID

            Try
                dsReturn = oFeeReceiptDB.DBGetDS(strSQL)

                If dsReturn.Tables(1).Rows.Count > 0 Then
                    dsRel1 = New DataRelation("CheckTransTOCheckTrans", dsReturn.Tables(0).Columns("Check_Trans_ID"), dsReturn.Tables(1).Columns("Check_Trans_ID"), False)
                    'dsRel2 = New DataRelation("CheckTransTypeTOCheckTransType", dsReturn.Tables(0).Columns("Trans_ID"), dsReturn.Tables(1).Columns("Trans_ID"), False)
                    dsReturn.Relations.Add(dsRel1)
                    'dsReturn.Relations.Add(dsRel2)

                End If
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        'Public Function MarkCalendarCompleted_ByDesc(ByVal strDesc As String) As String
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_Type = 25 AND Task_Description like '%" & strDesc & "%'"
        '        oFeeReceiptDB.DBExecNonQuery(strSQL)

        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        'Public Function GetMaxFeeReceiptID() As Int64
        '    Dim dsRemSys As New DataSet
        '    Dim dtReturn As Int64
        '    Dim strSQL As String

        '    Try

        '        strSQL = "select  * "
        '        strSQL &= "from vFeeReceiptGrid "
        '        strSQL &= "order by Fiscal_Year Desc, Fees_Basis_ID DESC "

        '        dsRemSys = oFeeReceiptDB.DBGetDS(strSQL)
        '        dtReturn = 0
        '        If dsRemSys.Tables(0).Rows.Count > 0 Then
        '            dtReturn = dsRemSys.Tables(0).Rows(0)("Fees_Basis_ID")
        '        End If
        '        Return dtReturn

        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        'Public Function PopulateFeeUnits() As DataTable
        '    Try
        '        Dim dtReturn As DataTable = GetDataTable("vFeeUnitsList")
        '        Return dtReturn
        '    Catch ex As Exception
        '        If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
        '        Throw ex
        '    End Try
        'End Function


        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal PropertyMaster As Boolean = True, Optional ByVal IncludeBlank As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            strSQL = ""
            If PropertyMaster Then
                If IncludeBlank Then
                    strSQL = " SELECT '' as PROPERTY_NAME, 0 as PROPERTY_ID, 0 as PROPERTY_POSITION "
                    strSQL &= " UNION "
                End If

                strSQL &= " SELECT PROPERTY_NAME, PROPERTY_ID, PROPERTY_POSITION FROM " & strProperty
                strSQL &= " order by 1 "
            Else
                strSQL &= " SELECT * FROM " & strProperty
            End If
            Try
                dsReturn = oFeeReceiptDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

#End Region



#End Region
#Region "External Event Handlers"

#End Region
    End Class
End Namespace

