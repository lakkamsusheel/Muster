Namespace MUSTER.BusinessLogic
    <Serializable()> Public Class pFinancial

#Region "Private Member Variables"
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        Private MusterException As Exceptions.MusterExceptions = New MUSTER.Exceptions.MusterExceptions
        Private colFinancial As MUSTER.Info.FinancialCollection = New MUSTER.Info.FinancialCollection
        Private oFinancialInfo As MUSTER.Info.FinancialInfo = New MUSTER.Info.FinancialInfo
        Private oFinancialDB As MUSTER.DataAccess.FinancialDB = New MUSTER.DataAccess.FinancialDB
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Financial").ID
        Private nID As Int64 = -1
#End Region
#Region "Exposed Events"
        Public Delegate Sub FinancialBLChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialBLColChangedEventHandler(ByVal bolValue As Boolean)
        Public Delegate Sub FinancialBLErrEventHandler(ByVal MsgStr As String)
        Public Delegate Sub FinancialInfoChanged()
        Public Event FinancialBLChanged As FinancialBLChangedEventHandler
        Public Event FinancialBLColChanged As FinancialBLColChangedEventHandler
        Public Event FinancialBLErr As FinancialBLErrEventHandler
        Public Event FinancialInfChanged As FinancialInfoChanged
        ' indicates change in the underlying _ProtoInfo structure
#End Region
#Region "Constructors"
        Public Sub New()
            oFinancialInfo = New MUSTER.Info.FinancialInfo
        End Sub
        Public Sub New(ByVal TextID As Integer)
            oFinancialInfo = New MUSTER.Info.FinancialInfo
            Me.Retrieve(TextID)
        End Sub

#End Region
#Region "Exposed Attributes"

        Public Property Sequence() As Integer
            Get
                Return oFinancialInfo.Sequence
            End Get
            Set(ByVal Value As Integer)
                oFinancialInfo.Sequence = Value
            End Set
        End Property

        Public Property TecEventID() As Int64
            Get
                Return oFinancialInfo.TecEventID
            End Get
            Set(ByVal Value As Int64)
                oFinancialInfo.TecEventID = Value
            End Set
        End Property
        Public ReadOnly Property TecEventIDDesc() As Int64
            Get
                Return oFinancialInfo.TecEventIDDesc
            End Get
        End Property
        Public Property StartDate() As Date
            Get
                Return oFinancialInfo.StartDate
            End Get
            Set(ByVal Value As Date)
                oFinancialInfo.StartDate = Value
            End Set
        End Property

        Public Property ClosedDate() As Date
            Get
                Return oFinancialInfo.ClosedDate
            End Get
            Set(ByVal Value As Date)
                oFinancialInfo.ClosedDate = Value
            End Set
        End Property

        Public Property VendorID() As Int64
            Get
                Return oFinancialInfo.VendorID
            End Get
            Set(ByVal Value As Int64)
                oFinancialInfo.VendorID = Value
            End Set
        End Property
        Public Property Status() As Int64
            Get
                Return oFinancialInfo.Status
            End Get
            Set(ByVal Value As Int64)
                oFinancialInfo.Status = Value
            End Set
        End Property

        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return oFinancialInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oFinancialInfo.CreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oFinancialInfo.CreatedOn
            End Get
        End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                Return oFinancialInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oFinancialInfo.Deleted = Value
            End Set
        End Property
        ' The entity ID associated with a Financial object.
        Public ReadOnly Property EntityID() As Integer
            Get
                Return oFinancialInfo.EntityID
            End Get
        End Property
        ' The unique ID for the row containing the text in the table (info.ID)
        Public ReadOnly Property ID() As Int64
            Get
                Return oFinancialInfo.ID
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oFinancialInfo.IsDirty
            End Get
            Set(ByVal Value As Boolean)
                oFinancialInfo.IsDirty = Value
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oFinancialInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oFinancialInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oFinancialInfo.ModifiedOn
            End Get
        End Property

#End Region
#Region "Exposed Operations"
#Region "General Operations"
        Public Sub Clear()
            oFinancialInfo = New MUSTER.Info.FinancialInfo
        End Sub
        Public Sub Reset()
            oFinancialInfo.Reset()
        End Sub
#End Region
#Region "Info Operations"
        Public Function Retrieve(ByVal EventID As Int64) As MUSTER.Info.FinancialInfo
            Try
                oFinancialInfo = oFinancialDB.DBGetByID(EventID)
                If oFinancialInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oFinancialInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function Retrieve_ByTechnicalEventID(ByVal EventID As Int64) As MUSTER.Info.FinancialInfo
            Try
                oFinancialInfo = oFinancialDB.DBGetByTechID(EventID)
                If oFinancialInfo.ID = 0 Then
                    nID -= 1
                End If

                Return oFinancialInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        ' Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, Optional ByVal strModuleName As String = "", Optional ByVal updateLastEditedBy As Boolean = True, Optional ByVal fromTechnical As Boolean = False)
            Try
                If Me.ValidateData(strModuleName) Then

                    oFinancialDB.Put(oFinancialInfo, moduleID, staffID, returnVal, updateLastEditedBy, fromTechnical)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If
                    oFinancialInfo.Archive()

                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Financial") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True

            Try
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"

                        ' if any validations failed
                        Exit Select
                End Select
                'If errStr.Length > 0 Or Not validateSuccess Then
                '    RaiseEvent LustEventErr(errStr)
                'End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#Region "Collection Operations"
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xFinancialInfo As MUSTER.Info.FinancialInfo
            Try
                For Each xFinancialInfo In colFinancial.Values
                    If xFinancialInfo.IsDirty Then
                        oFinancialInfo = xFinancialInfo
                        Me.Save(moduleID, staffID, returnVal)
                    End If
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Add(ByRef oFinancialInfoLocal As MUSTER.Info.FinancialInfo)
            Try
                colFinancial.Add(oFinancialInfoLocal)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        ' Removes the entity supplied from the collection
        Public Sub Remove(ByVal oFinancialInfoLocal As Object)
            Try
                colFinancial.Remove(oFinancialInfoLocal)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Sub
        ' Gets all the info
        Public Function GetAllByFacility(ByVal nFacility As Int64) As MUSTER.Info.FinancialCollection
            Try
                colFinancial.Clear()
                colFinancial = oFinancialDB.DBGetByFacility(nFacility)
                Return colFinancial
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function


#End Region
#Region " Populate Routines "
        Public Function GetCommitmentBalance(ByVal CommitmentID As Int64) As Double
            Dim dsRemSys As New DataSet
            Dim dtReturn As Double
            Dim strSQL As String

            Try

                strSQL = "select Balance "
                strSQL &= "from vFinancialCommitment_Grid "
                strSQL &= "where CommitmentID =  " & CommitmentID

                dsRemSys = oFinancialDB.DBGetDS(strSQL)
                dtReturn = 0
                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0).Rows(0)("Balance")
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateVendorList() As DataTable
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try
                strSQL = "SELECT * from vFinancialVendors Order By VendorName"

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateRolloversZeroesList() As DataSet
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try
                strSQL = " SELECT * from vFinancialRolloversZeroes2 "

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateRolloverForNewPO() As DataSet
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try
                strSQL = "select * from vFinancialRolloversForNewPO where Balance <> 0.00"

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateCommitmentTecDocList(ByVal EventID As Int64, ByVal CommitmentID As Int64, ByVal ForInvoice As Boolean, ByVal Mode As String) As DataSet
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try
                strSQL = String.Format(" exec spGetFinTechReports {0}, {1} ", EventID, CommitmentID)
                'strSQL &= " from dbo.tblTEC_EVENT_ACTIVITY_DOCUMENT  "
                'strSQL &= " where Deleted = 0 and Document_Property_ID in (Select Document_ID from vFinancialInvoicedDocuments)  "
                'strSQL &= " and Event_ID =  " & EventID & " "

                '---------------------------------------------------------------------------------------
                'If ForInvoice Then
                'strSQL &= " and Commitment_ID in ("
                'strSQL &= CommitmentID & ")"
                'Else
                '   If Mode = "ADD" Then
                'strSQL &= " and ((Commitment_ID = 0 and Due_Date is NULL) or (Commitment_ID = "
                'strSQL &= CommitmentID & " and  Date_Closed is null ))"
                'strSQL &= " and (Commitment_ID = 0 and Due_Date is NULL)"
                '   Else
                        'strSQL &= " and ((Commitment_ID = 0 and Due_Date is NULL) or (Commitment_ID = "
                        'strSQL &= CommitmentID & "))"
                '      strSQL &= " and Commitment_ID in ("
                '     strSQL &= CommitmentID & ")"
                ' End If
                ' End If
                '----------------------------------------------------------------------------------------

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateInvoiceDeductionReasons() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialDeductionReasons", False, False)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateInvoiceReimbursementRequestsList2(ByVal Reimbursement_ID As Int64) As DataTable
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try

                strSQL = "select Reimbursement_ID, convert(varchar,Received_Date,1) + ' - ' + Convert(varchar,Requested_Amount,1) as ReimbursementDesc "
                strSQL &= "from vFinancialPaymentReimbursementGrid "
                strSQL &= "where Reimbursement_ID =  " & Reimbursement_ID

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0)
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateInvoiceCommitmentList2(ByVal CommitmentID As Int64) As DataTable
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try

                strSQL = "select CommitmentID, convert(varchar,ApprovedDate,1) + ' - PO#' + PONumber + ' - ' + Task + ' - ' + Convert(varchar,(Commitment),1) + '/' + Convert(varchar,Balance,1) as CommitmentDesc "
                strSQL &= "from vFinancialCommitment_Grid "
                strSQL &= "where CommitmentID =  " & CommitmentID

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0)
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateInvoiceReimbursementRequestsList(ByVal EventID As Int64) As DataTable
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try


                'SELECT Reimbursement_ID,ReimbursementDesc from 
                '(
                'select Reimbursement_ID, convert(varchar,Received_Date,1) + ' - ' + Convert(varchar,Requested_Amount,1) as ReimbursementDesc, 
                '(select count(*) from tblfin_invoices where reimbursement_id = vFinancialPaymentReimbursementGrid.Reimbursement_ID and Final = 1) as FinalCount 
                'from vFinancialPaymentReimbursementGrid 
                'where Fin_Event_ID =  1016 and rawRequestedAmount > isnull(rawPaidAmount ,0) and INCOMPLETE = 0 
                ') as tmpView
                'where FinalCount <= 0


                strSQL = "SELECT Reimbursement_ID,ReimbursementDesc FROM ("
                strSQL &= "select Reimbursement_ID, convert(varchar,Received_Date,1) + ' - ' + Convert(varchar,Requested_Amount,1) as ReimbursementDesc, "
                strSQL &= " (select count(*) from tblfin_invoices where reimbursement_id = vFinancialPaymentReimbursementGrid.Reimbursement_ID and Final = 1) as FinalCount "
                strSQL &= "from vFinancialPaymentReimbursementGrid "
                strSQL &= "where Fin_Event_ID =  " & EventID
                strSQL &= " and rawRequestedAmount <> isnull(rawPaidAmount ,0)"
                strSQL &= " and INCOMPLETE = 0 "
                strSQL &= ") as tmpView WHERE FinalCount <= 0 "

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0)
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function PopulateInvoiceCommitmentList(ByVal EventID As Int64) As DataTable
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try

                strSQL = "select CommitmentID, convert(varchar,ApprovedDate,1) + ' - PO#' + PONumber + ' - ' + Task + ' - ' + Convert(varchar,(Commitment),1) + '/' + Convert(varchar,Balance,1) as CommitmentDesc "
                strSQL &= " from vFinancialCommitment_Grid "
                strSQL &= " where Fin_Event_ID =  " & EventID
                strSQL &= " and convert(varchar,Balance,1) <> '0.00' "
                'strSQL &= " and PONumber > '' "
                strSQL &= " and CommitmentID not in (SELECT CommitmentID from vFinancialRolloversForNewPO where len(NewPO) = 0 and rollover = 1)"

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0)
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function GetCommitmentDropDown(ByVal nFIN_EVENT_ID As Integer) As DataTable
            Dim dsReturn As New DataSet
            Dim sReturn As Single
            Dim strSQL As String
            Try
                strSQL = "Select  commitmentID,  Funding_Type + '      Date: ' + convert(nvarchar(12),Approveddate,9) + '   $' + convert(nvarchar(100),Commitment) + '     PO: ' + convert(nvarchar(30), PONumber) + '   ' + convert(nvarchar(300), Comments) as [DESC]  from vFinancialCommitment_Grid where fin_event_id = " & nFIN_EVENT_ID.ToString
                dsReturn = oFinancialDB.DBGetDS(strSQL)

                If dsReturn Is Nothing OrElse dsReturn.Tables.Count = 0 Then
                    Return Nothing
                End If

                Return dsReturn.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try

        End Function


        Public Function CommitmentGridDataset(ByVal showOpenOnly As Boolean) As DataSet
            Dim dsRemSys As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel1 As DataRelation
            Dim dsRel2 As DataRelation
            Dim strSQL As String
            Try


                'strSQL = "select * from vFinancialCommitment_Grid where Fin_Event_ID = " & Me.ID & " order by CommitmentID;"
                strSQL = "select * from vFinancialCommitment_Grid where Fin_Event_ID = " & Me.ID & " "
                If showOpenOnly Then
                    strSQL = strSQL & "and Balance <> '0.00' and len(Balance) > 0" & " "
                End If
                strSQL = strSQL & "order by CommitmentID;"
                strSQL = strSQL & " "
                strSQL = strSQL & "select * from vFinancialPastInvoice_Grid where Fin_Event_ID = " & Me.ID & ";"
                strSQL = strSQL & " "
                strSQL = strSQL & "select * from vFinancialCommitmentAdjustment_Grid where Fin_Event_ID = " & Me.ID & ";"

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(1).Rows.Count > 0 Then
                    dsRel1 = New DataRelation("CommitmentToInvoice", dsRemSys.Tables(0).Columns("CommitmentID"), dsRemSys.Tables(1).Columns("CommitmentID"), False)
                    dsRemSys.Relations.Add(dsRel1)
                End If

                If dsRemSys.Tables(2).Rows.Count > 0 Then
                    dsRel2 = New DataRelation("CommitmentToAdjustment", dsRemSys.Tables(0).Columns("CommitmentID"), dsRemSys.Tables(2).Columns("CommitmentID"), False)
                    dsRemSys.Relations.Add(dsRel2)
                End If

                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function CommitmentGridInvoicesOnlyDataset(ByVal CommitmentID As Int64) As DataSet
            Dim dsRemSys As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel1 As DataRelation
            Dim dsRel2 As DataRelation
            Dim strSQL As String
            Try

                strSQL = "select * from vFinancialPastInvoice_Grid where CommitmentID = " & CommitmentID & ""

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function CommitmentGridAdjustmentsOnlyDataset(ByVal CommitmentID As Int64) As DataSet
            Dim dsRemSys As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel1 As DataRelation
            Dim dsRel2 As DataRelation
            Dim strSQL As String
            Try

                strSQL = "select * from vFinancialCommitmentAdjustment_Grid where CommitmentID = " & CommitmentID & ""

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function CommitmentTotalsDatatable(ByVal GridType As Int16, ByVal showOpenOnly As Boolean, ByVal excludeFed As Boolean) As DataTable
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try

                strSQL = "Select Fin_Event_ID, Convert(varchar,EventCommitmentTotal,1) as EventCommitmentTotal, Convert(varchar,EventAdjustmentTotal,1) as EventAdjustmentTotal, "
                strSQL &= "Convert(varchar,EventPaymentTotal,1) as EventPaymentTotal, Convert(varchar,EventBalanceTotal,1) as EventBalanceTotal,  "
                strSQL &= "Convert(varchar,EventCommitmentTotal + EventAdjustmentTotal,1) as TotalCommitted from  "
                strSQL &= "(select Fin_Event_ID, sum(cast(Commitment as money)) as EventCommitmentTotal, sum(cast(adjustment as money)) as EventAdjustmentTotal, sum(cast(Payment as money)) as EventPaymentTotal, sum(cast(Balance as money))  as EventBalanceTotal "
                Select Case GridType
                    Case 0
                        strSQL &= "from vFinancialCommitment_Grid "
                    Case 1
                        strSQL &= "from vFinancialCommitment_Non3rd_Grid "
                    Case 2
                        strSQL &= "from vFinancialCommitment_3rdOnly_Grid "
                End Select
                strSQL &= "where Fin_Event_ID = " & Me.ID & " "
                If excludeFed Then
                    strSQL &= "and Funding_Type <> 'FED' "
                End If
                If showOpenOnly Then
                    strSQL &= "and Balance <> '0.00' and len(Balance) > 0" & " "
                End If
                strSQL &= "group by Fin_Event_ID) as tmpView "

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0)
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PaymentGridDataset(ByVal showOpenOnly As Boolean) As DataSet
            Dim dsRemSys As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel1 As DataRelation
            Dim strSQL As String
            Try


                strSQL = "select * from vFinancialPaymentReimbursementGrid where Fin_Event_ID = " & Me.ID & " "
                If showOpenOnly Then
                    strSQL = strSQL & "and commitment_id in (select commitmentid from vFinancialCommitment_Grid where Fin_Event_ID = " & Me.ID & " "
                    strSQL = strSQL & "and Balance <> '0.00' and len(Balance) > 0) "
                    'Added to include unpaid requests.
                    strSQL = strSQL & " union select * from vFinancialPaymentReimbursementGrid where Fin_Event_ID = " & Me.ID & " "
                    strSQL = strSQL & "and Commitment_Id = 0" & ";"
                Else
                    strSQL = strSQL & ";"
                End If
                strSQL = strSQL & "select * from vFinancialPaymentInvoiceGrid where Fin_Event_ID = " & Me.ID & " "
                If showOpenOnly Then
                    strSQL = strSQL & "and commitmentid in (select commitmentid from vFinancialCommitment_Grid where Fin_Event_ID = " & Me.ID & " "
                    strSQL = strSQL & "and Balance <> '0.00' and len(Balance) > 0)" & ";"
                Else
                    strSQL = strSQL & ";"
                End If

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(1).Rows.Count > 0 Then
                    dsRel1 = New DataRelation("ReimbursementToInvoice", dsRemSys.Tables(0).Columns("REIMBURSEMENT_ID"), dsRemSys.Tables(1).Columns("REIMBURSEMENT_ID"), False)
                    dsRemSys.Relations.Add(dsRel1)
                End If


                Return dsRemSys

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetMaxPaymentNumber(ByVal EventID As Int64) As Integer
            Dim ds As DataSet
            Dim strSQL As String
            Dim nMaxPaymentNum As Integer = 0
            Try
                strSQL = "SELECT ISNULL(MAX(ISNULL(PAYMENT_NUMBER,0)),0) AS PAYMENT_NUMBER " + _
                            "FROM tblFIN_REIMBURSEMENTS WHERE DELETED = 0 " + _
                            "AND FINANCIAL_EVENTID = " + EventID.ToString
                ds = oFinancialDB.DBGetDS(strSQL)

                If ds.Tables(0).Rows.Count > 0 Then
                    nMaxPaymentNum = ds.Tables(0).Rows(0)(0)
                End If
                Return nMaxPaymentNum
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PaymentTotalsDatatable(ByVal GridType As Int16, ByVal showOpenOnly As Boolean) As DataTable
            Dim dsRemSys As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String

            Try

                strSQL = "select Fin_Event_ID, CONVERT(varchar,sum(cast(Requested_Amount as Money)),1) as Requested_Amount,   CONVERT(varchar,sum(cast(Requested_Invoiced as Money)),1) as Requested_Invoiced,  CONVERT(varchar,sum(cast(Paid as Money)),1) as Paid "
                strSQL &= " from vFinancialPaymentReimbursementGrid "
                strSQL &= " Where Fin_Event_ID = " & Me.ID & " "
                If showOpenOnly Then
                    strSQL &= " and (commitment_id in (select commitmentid from vFinancialCommitment_Grid where Fin_Event_ID = " & Me.ID & " "
                    strSQL &= "and Balance <> '0.00' and len(Balance) > 0) or commitment_id = 0)" & " "
                End If
                strSQL &= " group by Fin_Event_ID "

                dsRemSys = oFinancialDB.DBGetDS(strSQL)

                If dsRemSys.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsRemSys.Tables(0)
                End If
                Return dtReturn

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function PopulateEligibleTecEvents(ByVal FacilityID As Integer) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String


            strSQL = "select * from vFinancialEligibleTechEvents where Facility_ID = " & FacilityID

            Try
                dsReturn = oFinancialDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try

        End Function

        Public Function PopulateFinancialStatus() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("vFinancialEventStatus")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
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
                dsReturn = oFinancialDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

#End Region
#Region "Miscellaneous Operations"
        Public Function GetProjectEngineer(ByVal nEventSequence As Integer, ByVal nFacilityID As Integer) As DataTable
            Dim dsSet As DataSet
            Try
                dsSet = oFinancialDB.GetProjectEngineer(nEventSequence, nFacilityID)
                Return dsSet.Tables(0)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
#End Region
#End Region
#Region "Private Operations"
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colFinancial.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colFinancial.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colFinancial.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region




    End Class
End Namespace
