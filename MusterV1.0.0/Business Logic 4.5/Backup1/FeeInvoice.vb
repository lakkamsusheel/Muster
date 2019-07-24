

Namespace MUSTER.BusinessLogic
    '  -------------------------------------------------------------------------------
    '  MUSTER.BusinessLogic.pFeeInvoice
    '  Provides the operations required to manipulate a Fee Invoice object.
    '  
    '  Copyright (C) 2004, 2005 CIBER, Inc.
    '  All rights reserved.
    '  
    '  Release   Initials    Date        Description
    '  1.0         AB       09/21/2005    Original class definition
    '  
    '  Function          Description
    '  -------------------------------------------------------------------------------
    '  Attribute          Description
    '  -------------------------------------------------------------------------------
    Public Class pFeeInvoice

#Region "Public Events"
        Public Event FeeInvoiceBLChanged As FeeInvoiceBLChangedEventHandler
        Public Event FeeInvoiceBLColChanged As FeeInvoiceBLColChangedEventHandler
        Public Event FeeInvoiceBLErr As FeeInvoiceBLErrEventHandler
        Public Event FeeInvoiceInfChanged As FeeInvoiceInfoChanged

        Public Delegate Sub FeeInvoiceBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeInvoiceBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FeeInvoiceBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FeeInvoiceInfoChanged()
#End Region
#Region "Private member variables"
        Private WithEvents oFeeInvoice As New MUSTER.Info.FeeInvoiceInfo
        Private WithEvents oFeeInvoiceCol As New MUSTER.Info.FeeInvoiceCollection
        Private oFeeInvoiceDB As New MUSTER.DataAccess.FeeInvoiceDB

        Private MusterExceptions As Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            oFeeInvoiceDB = New MUSTER.dataaccess.FeeInvoiceDB
            oFeeInvoiceCol = New MUSTER.Info.FeeInvoiceCollection
        End Sub
        Public Sub New(ByVal strName As String)
            oFeeInvoiceDB = New MUSTER.dataaccess.FeeInvoiceDB
            oFeeInvoiceCol = New MUSTER.Info.FeeInvoiceCollection
        End Sub
        Public Sub New(ByVal FeeInvoiceID As Integer)
            oFeeInvoiceDB = New MUSTER.dataaccess.FeeInvoiceDB
            oFeeInvoiceCol = New MUSTER.Info.FeeInvoiceCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property CheckNumber() As String
            Get
                Return oFeeInvoice.CheckNumber
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.CheckNumber = Value
            End Set
        End Property
        Public Property CheckTransID() As String
            Get
                Return oFeeInvoice.CheckTransID
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.CheckTransID = Value
            End Set
        End Property
        Public Property IssueZip() As String
            Get
                Return oFeeInvoice.IssueZip
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.IssueZip = Value
            End Set
        End Property
        Public Property IssueState() As String
            Get
                Return oFeeInvoice.IssueState
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.IssueState = Value
            End Set
        End Property
        Public Property IssueCity() As String
            Get
                Return oFeeInvoice.IssueCity
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.IssueCity = Value
            End Set
        End Property
        Public Property IssueAddr2() As String
            Get
                Return oFeeInvoice.IssueAddr2
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.IssueAddr2 = Value
            End Set
        End Property
        Public Property IssueAddr1() As String
            Get
                Return oFeeInvoice.IssueAddr1
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.IssueAddr1 = Value
            End Set
        End Property
        Public Property IssueName() As String
            Get
                Return oFeeInvoice.IssueName
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.IssueName = Value
            End Set
        End Property
        Public Property InvoiceType() As String
            Get
                Return oFeeInvoice.InvoiceType
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.InvoiceType = Value
            End Set
        End Property
        Public Property FiscalYear() As Int32
            Get
                Return oFeeInvoice.FiscalYear
            End Get
            Set(ByVal Value As Int32)
                oFeeInvoice.FiscalYear = Value
            End Set
        End Property

        Public ReadOnly Property ID() As Int64
            Get
                Return oFeeInvoice.ID
            End Get
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return oFeeInvoice.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFeeInvoice.IsDirty = Value
            End Set
        End Property

        Public Property Deleted() As Boolean
            Get
                Return oFeeInvoice.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFeeInvoice.Deleted = Value
            End Set
        End Property

        Public Property Description() As String
            Get
                Return oFeeInvoice.Description
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.Description = Value
            End Set
        End Property


        Public Property RecType() As String
            Get
                Return oFeeInvoice.RecType
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.RecType = Value
            End Set
        End Property
        Public Property FeeType() As String
            Get
                Return oFeeInvoice.FeeType
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.FeeType = Value
            End Set
        End Property

        Public Property InvoiceAdviceID() As String
            Get
                Return oFeeInvoice.InvoiceAdviceID
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.InvoiceAdviceID = Value
            End Set
        End Property
        Public Property OwnerID() As Int64
            Get
                Return oFeeInvoice.OwnerID
            End Get
            Set(ByVal Value As Int64)
                oFeeInvoice.OwnerID = Value
            End Set
        End Property

        Public Property InvoiceAmount() As Single
            Get
                Return oFeeInvoice.InvoiceAmount
            End Get
            Set(ByVal Value As Single)
                oFeeInvoice.InvoiceAmount = Value
            End Set
        End Property
        Public Property InvoiceLineAmount() As Single
            Get
                Return oFeeInvoice.InvoiceLineAmount
            End Get
            Set(ByVal Value As Single)
                oFeeInvoice.InvoiceLineAmount = Value
            End Set
        End Property


        Public Property WarrantNumber() As String
            Get
                Return oFeeInvoice.WarrantNumber
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.WarrantNumber = Value
            End Set
        End Property

        Public Property FacilityID() As Int64
            Get
                Return oFeeInvoice.FacilityID
            End Get
            Set(ByVal Value As Int64)
                oFeeInvoice.FacilityID = Value
            End Set
        End Property

        Public Property SequenceNumber() As Int16
            Get
                Return oFeeInvoice.SequenceNumber
            End Get
            Set(ByVal Value As Int16)
                oFeeInvoice.SequenceNumber = Value
            End Set
        End Property
        Public Property UnitPrice() As Single
            Get
                Return oFeeInvoice.UnitPrice
            End Get
            Set(ByVal Value As Single)
                oFeeInvoice.UnitPrice = Value
            End Set
        End Property


        Public Property Quantity() As String
            Get
                Return oFeeInvoice.Quantity
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.Quantity = Value
            End Set
        End Property
        Public Property WarrantDate() As Date
            Get
                Return oFeeInvoice.WarrantDate
            End Get
            Set(ByVal Value As Date)
                oFeeInvoice.WarrantDate = Value
            End Set
        End Property
        Public Property DueDate() As Date
            Get
                Return oFeeInvoice.DueDate
            End Get
            Set(ByVal Value As Date)
                oFeeInvoice.DueDate = Value
            End Set
        End Property
        Public Property CreditApplyTo() As String
            Get
                Return oFeeInvoice.CreditApplyTo
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.CreditApplyTo = Value
            End Set
        End Property

        Public Property TypeGeneration() As Boolean
            Get
                Return oFeeInvoice.TypeGeneration
            End Get
            Set(ByVal Value As Boolean)
                oFeeInvoice.TypeGeneration = Value
            End Set
        End Property
        Public Property Processed() As Boolean
            Get
                Return oFeeInvoice.Processed
            End Get
            Set(ByVal Value As Boolean)
                oFeeInvoice.Processed = Value
            End Set
        End Property
        Public Property InvoiceLineItems() As MUSTER.Info.FeeInvoiceCollection
            Get
                Return oFeeInvoice.InvoiceLineItems
            End Get
            Set(ByVal Value As MUSTER.Info.FeeInvoiceCollection)
                oFeeInvoice.InvoiceLineItems = Value
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return oFeeInvoice.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFeeInvoice.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFeeInvoice.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFeeInvoice.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFeeInvoice.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        ' Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal FeeInvoiceID As Int64) As MUSTER.Info.FeeInvoiceInfo
            Dim oFeeInvoiceInfoLocal As MUSTER.Info.FeeInvoiceInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oFeeInvoiceInfoLocal In oFeeInvoiceCol.Values
                    If oFeeInvoiceInfoLocal.ID = ID Then
                        If oFeeInvoiceInfoLocal.IsAgedData = True And oFeeInvoiceInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oFeeInvoice = oFeeInvoiceInfoLocal
                            Return oFeeInvoice
                        End If
                    End If
                Next
                If bolDataAged Then
                    oFeeInvoiceCol.Remove(oFeeInvoiceInfoLocal)
                End If
                oFeeInvoice = oFeeInvoiceDB.DBGetByID(FeeInvoiceID)
                oFeeInvoiceCol.Add(oFeeInvoice)
                Return oFeeInvoice
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        ' Obtains and returns a collection of line item invoices
        Public Function RetrieveLineItems(ByVal headerFeeInvoiceID As Int64, Optional ByVal showDeleted As Boolean = False) As MUSTER.Info.FeeInvoiceCollection
            Try
                Return oFeeInvoiceDB.DBGetLineItemsByHeaderInvoiceID(headerFeeInvoiceID, showDeleted)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        ' Obtains and returns an entity as called for by ID
        Public Function GetByFiscalYear(ByVal FeeInvoiceYear As Int32) As MUSTER.Info.FeeInvoiceInfo
            Dim oFeeInvoiceInfoLocal As MUSTER.Info.FeeInvoiceInfo
            Dim bolDataAged As Boolean = False
            Try
                For Each oFeeInvoiceInfoLocal In oFeeInvoiceCol.Values
                    If oFeeInvoiceInfoLocal.FiscalYear = FeeInvoiceYear Then
                        If oFeeInvoiceInfoLocal.IsAgedData = True And oFeeInvoiceInfoLocal.IsDirty = False Then
                            bolDataAged = True
                            Exit For
                        Else
                            oFeeInvoice = oFeeInvoiceInfoLocal
                            Return oFeeInvoice
                        End If
                    End If
                Next
                If bolDataAged Then
                    oFeeInvoiceCol.Remove(oFeeInvoiceInfoLocal)
                End If
                oFeeInvoice = oFeeInvoiceDB.DBGetByFiscalYear(FeeInvoiceYear)
                oFeeInvoiceCol.Add(oFeeInvoice)
                Return oFeeInvoice
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
                    oFeeInvoiceDB.put(oFeeInvoice, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFeeInvoice.IsDirty = False
                    oFeeInvoice.Archive()
                    RaiseEvent FeeInvoiceBLChanged(oFeeInvoice.IsDirty)
                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Public Sub SaveNewInvoice(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal bolAddRegActivityFee As Boolean = False)
            Dim xFeeInvoice As New MUSTER.Info.FeeInvoiceInfo
            Try
                If ValidateData() Then
                    oFeeInvoiceDB.put(oFeeInvoice, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFeeInvoice.IsDirty = False
                    oFeeInvoice.Archive()
                    'RaiseEvent FeeInvoiceBLChanged(oFeeInvoice.IsDirty)

                    For Each xFeeInvoice In oFeeInvoice.InvoiceLineItems.Values
                        xFeeInvoice.InvoiceAdviceID = oFeeInvoice.InvoiceAdviceID
                        If xFeeInvoice.ID <= 0 Then
                            xFeeInvoice.CreatedBy = oFeeInvoice.CreatedBy
                        Else
                            xFeeInvoice.ModifiedBy = oFeeInvoice.ModifiedBy
                        End If
                        oFeeInvoiceDB.put(xFeeInvoice, moduleID, staffID, returnVal, bolAddRegActivityFee)
                        If Not returnVal = String.Empty Then
                            Exit Sub
                        End If

                        xFeeInvoice.IsDirty = False
                        xFeeInvoice.Archive()
                    Next
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
                ''oFeeInvoice = oFeeInvoiceDB.DBGetByID(ID)
                oFeeInvoice.ID = ID
                If oFeeInvoice.ID = 0 Then
                    'oFeeInvoice.ID = nID
                    'nID -= 1
                End If
                oFeeInvoiceCol.Add(oFeeInvoice)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef FeeInvoice As MUSTER.Info.FeeInvoiceInfo)
            Try
                oFeeInvoice = FeeInvoice
                'oFeeInvoice.UserID = onUserID
                oFeeInvoiceCol.Add(oFeeInvoice)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oFeeInvoiceLocal As MUSTER.Info.FeeInvoiceInfo

            Try
                For Each oFeeInvoiceLocal In oFeeInvoiceCol.Values
                    If oFeeInvoiceLocal.ID = ID Then
                        oFeeInvoiceCol.Remove(oFeeInvoiceLocal)
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
        Public Sub Remove(ByVal oLustEvent As MUSTER.Info.FeeInvoiceInfo)
            Try
                oFeeInvoiceCol.Remove(oLustEvent)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LustEvent " & oLustEvent.ID & " is not in the collection of FeeInvoice.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String)
            Dim xFeeInvoiceInfo As MUSTER.Info.FeeInvoiceInfo
            For Each xFeeInvoiceInfo In oFeeInvoiceCol.Values
                If xFeeInvoiceInfo.IsDirty Then
                    oFeeInvoice = xFeeInvoiceInfo
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
                    If oFeeInvoice.ID <= 0 Then
                        oFeeInvoice.CreatedBy = UserID
                    Else
                        oFeeInvoice.ModifiedBy = UserID
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
            '    Dim strArr() As String = oFeeInvoiceCol.GetKeys()
            '    Dim nArr(strArr.GetUpperBound(0)) As Integer
            '    Dim y As String
            '    For Each y In strArr
            '        nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            '    Next
            '    nArr.Sort(nArr)
            '    colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            '    If colIndex + direction > -1 And _
            '        colIndex + direction <= nArr.GetUpperBound(0) Then
            '        Return oFeeInvoiceCol.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            '    Else
            '        Return oFeeInvoiceCol.Item(nArr.GetValue(colIndex)).ID.ToString
            '    End If
        End Function
#End Region
#Region "General Operations"

        Public Function Reset() As Object
            oFeeInvoice.Reset()
            oFeeInvoiceCol.Clear()
        End Function
#End Region
#Region "Miscellaneous Operations"

#End Region

#Region " Populate Routines "
        'Public Function ApproveBilling() As String
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "UPDATE dbo.tblFEES_PENDING_INVOICES SET PROCESSED = 1 "
        '        oFeeInvoiceDB.DBExecNonQuery(strSQL)


        '        strSQL = "EXEC spFeesInsertToPermInvBP2KTemp"
        '        oFeeInvoiceDB.DBExecNonQuery(strSQL)

        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function
        '
        'Public Function MarkCalendarCompleted_ByDesc(ByVal strDesc As String) As String
        '    Dim strReturn As String
        '    Dim strSQL As String
        '    Try
        '        strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_Type = 25 AND Task_Description like '%" & strDesc & "%'"
        '        oFeeInvoiceDB.DBExecNonQuery(strSQL)

        '    Catch Ex As Exception
        '        If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
        '        Throw Ex
        '    End Try
        'End Function

        Public Function GetFacilityOldAccessInfo_ByFacility(ByVal FacilityID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim dsRel1 As DataRelation
            Dim strSQL As String


            strSQL = "select FY as FiscalYear, TankID, TransNum, (case When isnull(Amount,0) < 0 Then CONVERT(varchar, Amount * -1, 1) ELSE '' end) as Debit, (case When isnull(Amount,0) > 0 Then convert(varchar,Amount,1) ELSE '' end) as Credit, 0.00 as Balance, [Transaction] as Trans, DateIssued, DateDue,  Comments, StampInit,  stampdate from tblFeeTransaction where FacilityID = " & FacilityID & " Order by stampdate, TankID, [Transaction]"

            Try
                dsReturn = oFeeInvoiceDB.DBGetDS(strSQL)

                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetCreditMemos_ByOwnerID(ByVal OwnerID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim strSQL As String
            Dim dsRel1 As DataRelation

            strSQL = "select CreditMemoInvoice, CreditMemoDate, BillingInvoice, CreditTotal, INV_ID as InvoiceID from vFees_OwnerCreditMemoGrid where Inv_Number is not null and Owner_ID = " & OwnerID & " and CreditMemoInvoice not in (Select Credit_Apply_To from tblFees_Invoices where Invoice_Type = 'D') ;"
            strSQL = strSQL & " "
            strSQL = strSQL & "select CreditMemoInvoice, Facility_ID, FacilityName, LineAmount, INV_ID as InvoiceID from vFees_OwnerCreditMemoLineItemGrid where Inv_Number is not null and Owner_ID = " & OwnerID & " and CreditMemoInvoice not in (Select Credit_Apply_To from tblFees_Invoices where Invoice_Type = 'D') "
            Try
                dsReturn = oFeeInvoiceDB.DBGetDS(strSQL)

                If dsReturn.Tables(1).Rows.Count > 0 Then
                    dsRel1 = New DataRelation("InvToInv", dsReturn.Tables(0).Columns("CreditMemoInvoice"), dsReturn.Tables(1).Columns("CreditMemoInvoice"), False)
                    dsReturn.Relations.Add(dsRel1)
                End If
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetRefunds_ByOwnerID(ByVal OwnerID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim strSQL As String
            Dim dsRel1 As DataRelation


            strSQL = "select * from vFees_OwnerRefundSummaryGrid where Owner_ID = " & OwnerID & ";"
            strSQL = strSQL & " "
            strSQL = strSQL & "select Date_Created as TransactionDate, Advice_ID, Change_Type, Warrant_Number, New_Warrant_Number, Warrant_date, New_Warrant_Date from vFees_OwnerRefundSummaryLineItemGrid where Owner_ID = " & OwnerID & " Order By Date_Created"

            Try
                dsReturn = oFeeInvoiceDB.DBGetDS(strSQL)


                If dsReturn.Tables(1).Rows.Count > 0 Then
                    dsRel1 = New DataRelation("AdvToAdv", dsReturn.Tables(0).Columns("Advice_ID"), dsReturn.Tables(1).Columns("Advice_ID"), False)
                    dsReturn.Relations.Add(dsRel1)
                End If

                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetInvoiceLineItemSummaryGrid_ByInvoiceID(ByVal InvoiceID As Int64, Optional ByVal bolUpdateCM As Boolean = False) As DataSet
            Dim strSQL As String

            If bolUpdateCM Then
                strSQL = "select " + _
                            "(select max(INV__ID) from tblFees_Invoices where deleted = 0 and fee_type = 'c' and rec_type = 'advic' and Credit_Apply_To IS NOT NULL AND Credit_Apply_To IN " + _
                                "(select INV_NUMBER from tblFees_Invoices where INV__ID = " + InvoiceID.ToString + ") " + _
                            "and DATE_CREATED > '" + Today.Date.ToString + "'" + _
                            ") as CM_INV__ID, " + _
                            "b.inv__id as CM_LINE_INV__ID, a.Facility_ID, a.FacilityName, a.Fiscal_Year, a.Charges, a.Balance, " + _
                            "CONVERT(varchar, ISNULL(b.INV_LINE_AMT,0.00), 1) as Credits, ISNULL(b.Description,a.LineDescription) as LineDescription, a.Quantity, a.UnitPrice, a.ITEM_SEQ_NUMBER from " + _
                            "vFees_OwnerInvoiceSummaryLineItemGrid a left outer join " + _
                                "(select * from tblFees_Invoices where deleted = 0 and advice_id in " + _
                                    "(select advice_id from tblFees_Invoices where deleted = 0 and fee_type = 'c' and rec_type = 'advic' and Credit_Apply_To IS NOT NULL AND Credit_Apply_To IN " + _
                                        "(select INV_NUMBER from tblFees_Invoices where INV__ID = " + InvoiceID.ToString + ") " + _
                                    " and DATE_CREATED > '" + Today.Date.ToString + "'" + _
                                    ")" + _
                                    " and rec_type = 'adln' and " + _
                                    "DATE_CREATED > '" + Today.Date.ToString + "'" + _
                                ") b " + _
                            "on a.facility_id = b.facility_id " + _
                            "where a.Advice_ID in (select Advice_ID from tblFees_Invoices where INV__ID = " + InvoiceID.ToString + ")"
            Else
                strSQL = "select NULL AS CM_INV__ID, NULL as CM_LINE_INV__ID, Facility_ID, FacilityName, Fiscal_Year, Charges, Balance, 0.00 as Credits, LineDescription, Quantity, UnitPrice, ITEM_SEQ_NUMBER from vFees_OwnerInvoiceSummaryLineItemGrid where Advice_ID in (select Advice_ID from tblFees_Invoices where INV__ID = " & InvoiceID & ")"
            End If
            Try
                Return oFeeInvoiceDB.DBGetDS(strSQL)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function GetInvoiceLineItemBalanceDue_ByInvoiceID(ByVal InvoiceID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim strSQL As String


            strSQL = "select Facility_ID, FacilityName, Fiscal_Year, Charges, Balance, 0.00 as Reallocation, '' as CheckNumber, '' as Reason, ITEM_SEQ_NUMBER from vFees_OwnerInvoiceSummaryLineItemGrid where Advice_ID in (select Advice_ID from tblFees_Invoices where INV__ID = " & InvoiceID & ") and Balance > '0.00'"

            Try
                dsReturn = oFeeInvoiceDB.DBGetDS(strSQL)

                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function



        Public Function GetFacilityFeeTransaction_ByFacility(ByVal FacilityID As Int64) As DataSet
            Dim strSQL As String
            'strSQL = "select * from (" + _
            '         "select FiscalYear, Inv_Number, INV_Date, Due_Date, TransactionType, Reference, Debit, Credit, " + _
            '         "balance, Description, Facility_ID, CreditAppliedTo, DATE_CREATED from vFEES_FacilityFeeByTransaction where Facility_ID = " + FacilityID.ToString + _
            '         " union all " + _
            '         "select dbo.udfgetfiscalyear(getdate()) - 1 as FiscalYear, " + _
            '         "'' AS Inv_Number, '07/01/' + cast(dbo.udfgetfiscalyear(getdate()) - 2 as varchar) as INV_Date, " + _
            '         "null as Due_Date, 'Prior Balance' AS TransactionType, '' AS Reference, " + _
            '         "convert(varchar, sum(cast(Debit as money)), 1) AS Debit, " + _
            '         "convert(varchar, sum(cast(Credit as money)), 1) AS Credit, " + _
            '         "convert(varchar, sum(cast(balance as money)), 1) AS balance, " + _
            '         "'' as Description, Facility_ID, NULL as CreditAppliedTo, null as DATE_CREATED " + _
            '         "from vFEES_FacilityFeeByTransaction " + _
            '         "where facility_id = " + FacilityID.ToString + " and FiscalYear < dbo.udfgetfiscalyear(getdate()) " + _
            '         "and TransactionType <> 'Prior Balance' group by facility_id) tmpV " + _
            '         "Order by FiscalYear, INV_Date"


            strSQL = String.Format("sp_Fees_facilityFeeTransactionPrioritziedSummary {0}", FacilityID.ToString)


            Try
                Return oFeeInvoiceDB.DBGetDS(strSQL)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetFacilityFeeInvoice_ByFacility(ByVal FacilityID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim dsRel1 As DataRelation
            Dim strSQL As String


            strSQL = "Select Inv_ID, FiscalYear, Inv_Number, Advice_ID, FeeTypeDesc, CONVERT(varchar, Inv_Line_Amt, 1) as Inv_Line_Amt, Balance "
            strSQL = strSQL & " from (select (Select Inv_ID from vFeesInvoiceRecords vFIR where vFIR.Advice_ID = tmpView.Advice_ID and Rec_Type = 'ADVIC') as Inv_ID, "
            strSQL = strSQL & " (Select SFY from vFeesInvoiceRecords vFIR where vFIR.Advice_ID = tmpView.Advice_ID and Rec_Type = 'ADVIC') as FiscalYear, "
            strSQL = strSQL & " isnull(Inv_Number,Advice_ID) as Inv_Number, Advice_ID, FeeTypeDesc, Inv_Line_Amt, '0.00' as Balance "
            strSQL = strSQL & " from(Select Inv_Number, Advice_ID, FeeTypeDesc, SUM(INV_Line_AMT) as Inv_Line_Amt "
            strSQL = strSQL & " from vFeesInvoiceRecords  "
            strSQL = strSQL & " where Invoice_Type = 'I' and Facility_ID =  " & FacilityID
            strSQL = strSQL & " Group By Inv_Number, Advice_ID, FeeTypeDesc) as tmpView) as tmpView2 "
            strSQL = strSQL & " Order By Inv_ID "
            strSQL = strSQL & "; "
            strSQL = strSQL & "Select FiscalYear, (Case isnull(CreditAppliedTo,'') When '' THEN Inv_Number ELSE (Case TransactionType WHEN 'Credit Memo' THEN ltrim(rtrim(CreditAppliedTo)) ELSE (Select top 1 vFI.Credit_Apply_To from tblFees_Invoices vFI where vFI.INV_Number = vFEES_FacilityFeeByTransaction.CreditAppliedTo) END )END) as Inv_Number, Inv_date, Due_Date, TransactionType, Reference, Debit, Credit, Balance, Description, Facility_ID, (Case isnull(CreditAppliedTo,'') When '' THEN '' ELSE Inv_Number END) as  CreditAppliedTo, DATE_CREATED from vFEES_FacilityFeeByTransaction where Facility_ID = " & FacilityID & " Order BY FiscalYear, INV_Date "

            Try
                dsReturn = oFeeInvoiceDB.DBGetDS(strSQL)

                If dsReturn.Tables(1).Rows.Count > 0 Then
                    dsRel1 = New DataRelation("Inv1ToInv2", dsReturn.Tables(0).Columns("Inv_Number"), dsReturn.Tables(1).Columns("Inv_Number"), False)
                    dsReturn.Relations.Add(dsRel1)
                End If
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetFacilitySummaryInvoiceGrid_ByOwnerID(ByVal OwnerID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim dsRel1 As DataRelation
            Dim strSQL As String


            strSQL = "select * from vFees_OwnerInvoiceSummaryGrid where Owner_ID = " & OwnerID & ";"
            strSQL = strSQL & " "
            strSQL = strSQL & "select * from vFees_OwnerInvoiceSummaryLineItemGrid where Owner_ID = " & OwnerID

            Try
                dsReturn = oFeeInvoiceDB.DBGetDS(strSQL)

                If dsReturn.Tables(1).Rows.Count > 0 Then
                    dsRel1 = New DataRelation("ADVICToADLN", dsReturn.Tables(0).Columns("InvoiceID"), dsReturn.Tables(1).Columns("InvoiceID"), False)
                    dsReturn.Relations.Add(dsRel1)
                End If
                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetFacilitySummaryGrid_ByOwnerID(ByVal OwnerID As Int64) As DataSet
            Dim dsReturn As New DataSet
            Dim strSQL As String

            strSQL = "SELECT * FROM vFees_FacilitySummaryGrid where Owner_ID = " & OwnerID & " Order by Facility_ID"

            Try
                dsReturn = oFeeInvoiceDB.DBGetDS(strSQL)

                Return dsReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetOverpaymentBucket(ByVal OwnerID As Int64) As Single
            Dim dsRemSys As New DataSet
            Dim dtReturn As Single
            Dim strSQL As String

            Try

                strSQL = "select OverpaymentBucket from vFees_Overpayment_Bucket where Owner_ID =  " & OwnerID

                dsRemSys = oFeeInvoiceDB.DBGetDS(strSQL)
                dtReturn = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0).Rows(0)("OverpaymentBucket")
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetCurrentBalance(ByVal OwnerID As Int64) As Single
            Dim dsRemSys As New DataSet
            Dim dtReturn As Single
            Dim strSQL As String

            Try

                strSQL = "select isnull(sum(cast(ToDateBalance as money)),0) as CurrentBalance from vFees_FacilitySummaryGrid where Owner_ID = " & OwnerID

                dsRemSys = oFeeInvoiceDB.DBGetDS(strSQL)
                dtReturn = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0).Rows(0)("CurrentBalance")
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetCurrentBalance_Facility(ByVal FacilityID As Int64) As Single
            Dim dsRemSys As New DataSet
            Dim returnVal As Single
            Dim strSQL As String

            Try

                strSQL = "select isnull(sum(cast(ToDateBalance as money)),0) as CurrentBalance from vFees_FacilitySummaryGrid where Facility_ID = " & FacilityID

                dsRemSys = oFeeInvoiceDB.DBGetDS(strSQL)
                returnVal = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    returnVal = dsRemSys.Tables(0).Rows(0)("CurrentBalance")
                End If
                Return returnVal
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetOwnerFeeBalanceOn(ByVal onDate As DateTime, Optional ByVal OwnerID As Int64 = 0, Optional ByVal FacilityID As Int64 = 0) As Single
            Dim dsRemSys As New DataSet
            Dim returnVal As Single
            Dim strSQL As String

            Try
                If Date.Compare(onDate, CDate("01/01/0001")) = 0 Then
                    strSQL = "exec spFeesGetOwnerFeeBalanceOn " + IIf(OwnerID = 0, "NULL", OwnerID.ToString) + ",  " + IIf(FacilityID = 0, "NULL", FacilityID.ToString) + ", NULL"
                Else
                    strSQL = "exec spFeesGetOwnerFeeBalanceOn " + IIf(OwnerID = 0, "NULL", OwnerID.ToString) + ",  " + IIf(FacilityID = 0, "NULL", FacilityID.ToString) + ", '" + onDate.ToShortDateString + "'"
                End If

                dsRemSys = oFeeInvoiceDB.DBGetDS(strSQL)
                returnVal = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    returnVal = dsRemSys.Tables(0).Rows(0)("Balance")
                End If
                Return returnVal
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetRefundableCheckAmount(ByVal OwnerID As Int64, ByVal CheckTransID As String, ByVal CheckNumber As String) As Single
            Dim dsRemSys As New DataSet
            Dim dtReturn As Int64
            Dim strSQL As String

            Try

                strSQL = "select  AvailableOverpayment "
                strSQL &= "from vFees_RefundableChecks "
                strSQL &= " Where Owner_ID = " & OwnerID
                strSQL &= " AND Check_trans_ID = '" & CheckTransID & "'"
                'strSQL &= " AND Check_Number = '" & CheckNumber & "'"

                dsRemSys = oFeeInvoiceDB.DBGetDS(strSQL)
                dtReturn = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0).Rows(0)("AvailableOverpayment")
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function GetCheckIssuerName(ByVal OwnerID As Int64, ByVal CheckTransID As String, ByVal CheckNumber As String) As String
            Dim dsRemSys As New DataSet
            Dim dtReturn As String
            Dim strSQL As String

            Try

                strSQL = "select top 1 Issuing_company  "
                strSQL &= " from tblFees_Receipts "
                strSQL &= " Where Owner_ID = " & OwnerID
                strSQL &= " AND Check_trans_ID = '" & CheckTransID & "'"
                'strSQL &= " AND Check_Number = '" & CheckNumber & "'"

                dsRemSys = oFeeInvoiceDB.DBGetDS(strSQL)
                dtReturn = ""
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0).Rows(0)("Issuing_company")
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GenerateDebitMemo(ByVal CreditMemoID As Int64, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String) As Boolean
            Try

                oFeeInvoiceDB.GenerateDebitMemo(CreditMemoID, moduleID, staffID, returnVal, UserID)
                If Not returnVal = String.Empty Then
                    Exit Function
                End If

                Return True
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterExceptions.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


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
                dsReturn = oFeeInvoiceDB.DBGetDS(strSQL)
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

