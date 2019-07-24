Public Class GenerateCreditMemo
    Inherits System.Windows.Forms.Form
#Region " User Defined Variables"
    Friend InvoiceID As Int64
    Friend OwnerID As Int64
    Friend FacilityID As Int64
    Friend bolUpdateCM As Boolean = False
    Friend CallingForm As Form

    'Private frmCreditMemoSummary As CreditMemoSummary
    Dim oInvoice As New MUSTER.BusinessLogic.pFeeInvoice
    Dim returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents pnlGenCreditMemoBottom As System.Windows.Forms.Panel
    Friend WithEvents btnIssueCreditMemo As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents pnlGenCreditMemoHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlGenCreditMemoDetails As System.Windows.Forms.Panel
    Friend WithEvents ugGenCreditMemo As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblOrigInvoice As System.Windows.Forms.Label
    Friend WithEvents lblOrigInvoiceValue As System.Windows.Forms.Label
    Friend WithEvents lblOriginalAmt As System.Windows.Forms.Label
    Friend WithEvents lblOriginalAmtValue As System.Windows.Forms.Label
    Friend WithEvents lblCreditAmount As System.Windows.Forms.Label
    Friend WithEvents lblCreditAmtValue As System.Windows.Forms.Label
    Friend WithEvents lblFeeType As System.Windows.Forms.Label
    Friend WithEvents lblFeeTypeValue As System.Windows.Forms.Label
    Friend WithEvents lblAdviceID As System.Windows.Forms.Label
    Friend WithEvents lblAdviceIDValue As System.Windows.Forms.Label
    Friend WithEvents btnDeleteCreditMemo As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlGenCreditMemoBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnIssueCreditMemo = New System.Windows.Forms.Button
        Me.btnDeleteCreditMemo = New System.Windows.Forms.Button
        Me.pnlGenCreditMemoHeader = New System.Windows.Forms.Panel
        Me.lblAdviceIDValue = New System.Windows.Forms.Label
        Me.lblAdviceID = New System.Windows.Forms.Label
        Me.lblFeeTypeValue = New System.Windows.Forms.Label
        Me.lblFeeType = New System.Windows.Forms.Label
        Me.lblCreditAmtValue = New System.Windows.Forms.Label
        Me.lblCreditAmount = New System.Windows.Forms.Label
        Me.lblOriginalAmtValue = New System.Windows.Forms.Label
        Me.lblOriginalAmt = New System.Windows.Forms.Label
        Me.lblOrigInvoiceValue = New System.Windows.Forms.Label
        Me.lblOrigInvoice = New System.Windows.Forms.Label
        Me.pnlGenCreditMemoDetails = New System.Windows.Forms.Panel
        Me.ugGenCreditMemo = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlGenCreditMemoBottom.SuspendLayout()
        Me.pnlGenCreditMemoHeader.SuspendLayout()
        Me.pnlGenCreditMemoDetails.SuspendLayout()
        CType(Me.ugGenCreditMemo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlGenCreditMemoBottom
        '
        Me.pnlGenCreditMemoBottom.Controls.Add(Me.btnCancel)
        Me.pnlGenCreditMemoBottom.Controls.Add(Me.btnIssueCreditMemo)
        Me.pnlGenCreditMemoBottom.Controls.Add(Me.btnDeleteCreditMemo)
        Me.pnlGenCreditMemoBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlGenCreditMemoBottom.Location = New System.Drawing.Point(0, 342)
        Me.pnlGenCreditMemoBottom.Name = "pnlGenCreditMemoBottom"
        Me.pnlGenCreditMemoBottom.Size = New System.Drawing.Size(688, 40)
        Me.pnlGenCreditMemoBottom.TabIndex = 2
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(348, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(112, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        '
        'btnIssueCreditMemo
        '
        Me.btnIssueCreditMemo.Location = New System.Drawing.Point(228, 8)
        Me.btnIssueCreditMemo.Name = "btnIssueCreditMemo"
        Me.btnIssueCreditMemo.Size = New System.Drawing.Size(112, 23)
        Me.btnIssueCreditMemo.TabIndex = 4
        Me.btnIssueCreditMemo.Text = "Issue Credit Memo"
        '
        'btnDeleteCreditMemo
        '
        Me.btnDeleteCreditMemo.Location = New System.Drawing.Point(96, 8)
        Me.btnDeleteCreditMemo.Name = "btnDeleteCreditMemo"
        Me.btnDeleteCreditMemo.Size = New System.Drawing.Size(120, 23)
        Me.btnDeleteCreditMemo.TabIndex = 4
        Me.btnDeleteCreditMemo.Text = "Delete Credit Memo"
        '
        'pnlGenCreditMemoHeader
        '
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblAdviceIDValue)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblAdviceID)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblFeeTypeValue)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblFeeType)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblCreditAmtValue)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblCreditAmount)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblOriginalAmtValue)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblOriginalAmt)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblOrigInvoiceValue)
        Me.pnlGenCreditMemoHeader.Controls.Add(Me.lblOrigInvoice)
        Me.pnlGenCreditMemoHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlGenCreditMemoHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlGenCreditMemoHeader.Name = "pnlGenCreditMemoHeader"
        Me.pnlGenCreditMemoHeader.Size = New System.Drawing.Size(688, 72)
        Me.pnlGenCreditMemoHeader.TabIndex = 2
        '
        'lblAdviceIDValue
        '
        Me.lblAdviceIDValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdviceIDValue.Location = New System.Drawing.Point(344, 16)
        Me.lblAdviceIDValue.Name = "lblAdviceIDValue"
        Me.lblAdviceIDValue.TabIndex = 10
        '
        'lblAdviceID
        '
        Me.lblAdviceID.Location = New System.Drawing.Point(248, 16)
        Me.lblAdviceID.Name = "lblAdviceID"
        Me.lblAdviceID.TabIndex = 9
        Me.lblAdviceID.Text = "Advice Number"
        '
        'lblFeeTypeValue
        '
        Me.lblFeeTypeValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFeeTypeValue.Location = New System.Drawing.Point(536, 16)
        Me.lblFeeTypeValue.Name = "lblFeeTypeValue"
        Me.lblFeeTypeValue.Size = New System.Drawing.Size(144, 23)
        Me.lblFeeTypeValue.TabIndex = 8
        '
        'lblFeeType
        '
        Me.lblFeeType.Location = New System.Drawing.Point(472, 16)
        Me.lblFeeType.Name = "lblFeeType"
        Me.lblFeeType.Size = New System.Drawing.Size(64, 23)
        Me.lblFeeType.TabIndex = 7
        Me.lblFeeType.Text = "Fee Type"
        '
        'lblCreditAmtValue
        '
        Me.lblCreditAmtValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCreditAmtValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCreditAmtValue.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.lblCreditAmtValue.Location = New System.Drawing.Point(344, 40)
        Me.lblCreditAmtValue.Name = "lblCreditAmtValue"
        Me.lblCreditAmtValue.TabIndex = 6
        Me.lblCreditAmtValue.Text = "0.00"
        Me.lblCreditAmtValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCreditAmount
        '
        Me.lblCreditAmount.Location = New System.Drawing.Point(248, 40)
        Me.lblCreditAmount.Name = "lblCreditAmount"
        Me.lblCreditAmount.TabIndex = 5
        Me.lblCreditAmount.Text = "Credit Amount"
        '
        'lblOriginalAmtValue
        '
        Me.lblOriginalAmtValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOriginalAmtValue.Location = New System.Drawing.Point(112, 40)
        Me.lblOriginalAmtValue.Name = "lblOriginalAmtValue"
        Me.lblOriginalAmtValue.TabIndex = 4
        Me.lblOriginalAmtValue.Text = "0.00"
        Me.lblOriginalAmtValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOriginalAmt
        '
        Me.lblOriginalAmt.Location = New System.Drawing.Point(8, 40)
        Me.lblOriginalAmt.Name = "lblOriginalAmt"
        Me.lblOriginalAmt.TabIndex = 3
        Me.lblOriginalAmt.Text = "Original Amount"
        '
        'lblOrigInvoiceValue
        '
        Me.lblOrigInvoiceValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOrigInvoiceValue.Location = New System.Drawing.Point(112, 16)
        Me.lblOrigInvoiceValue.Name = "lblOrigInvoiceValue"
        Me.lblOrigInvoiceValue.TabIndex = 2
        '
        'lblOrigInvoice
        '
        Me.lblOrigInvoice.Location = New System.Drawing.Point(8, 16)
        Me.lblOrigInvoice.Name = "lblOrigInvoice"
        Me.lblOrigInvoice.TabIndex = 1
        Me.lblOrigInvoice.Text = "Invoice"
        '
        'pnlGenCreditMemoDetails
        '
        Me.pnlGenCreditMemoDetails.Controls.Add(Me.ugGenCreditMemo)
        Me.pnlGenCreditMemoDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlGenCreditMemoDetails.Location = New System.Drawing.Point(0, 72)
        Me.pnlGenCreditMemoDetails.Name = "pnlGenCreditMemoDetails"
        Me.pnlGenCreditMemoDetails.Size = New System.Drawing.Size(688, 270)
        Me.pnlGenCreditMemoDetails.TabIndex = 0
        '
        'ugGenCreditMemo
        '
        Me.ugGenCreditMemo.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugGenCreditMemo.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugGenCreditMemo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugGenCreditMemo.Location = New System.Drawing.Point(0, 0)
        Me.ugGenCreditMemo.Name = "ugGenCreditMemo"
        Me.ugGenCreditMemo.Size = New System.Drawing.Size(688, 270)
        Me.ugGenCreditMemo.TabIndex = 1
        '
        'GenerateCreditMemo
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(688, 382)
        Me.Controls.Add(Me.pnlGenCreditMemoDetails)
        Me.Controls.Add(Me.pnlGenCreditMemoHeader)
        Me.Controls.Add(Me.pnlGenCreditMemoBottom)
        Me.Name = "GenerateCreditMemo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Generate Credit Memo"
        Me.pnlGenCreditMemoBottom.ResumeLayout(False)
        Me.pnlGenCreditMemoHeader.ResumeLayout(False)
        Me.pnlGenCreditMemoDetails.ResumeLayout(False)
        CType(Me.ugGenCreditMemo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region " Form Load Events"
    Private Sub GenerateCreditMemo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oInvoice.Retrieve(InvoiceID)
        lblOrigInvoiceValue.Text = oInvoice.WarrantNumber
        lblOriginalAmtValue.Text = FormatNumber(oInvoice.InvoiceAmount, 2, TriState.True, TriState.False, TriState.True)
        lblAdviceIDValue.Text = oInvoice.InvoiceAdviceID
        If oInvoice.FeeType = "FD" Then
            lblFeeTypeValue.Text = "Annual Main"
        ElseIf oInvoice.FeeType = "1" Then
            lblFeeTypeValue.Text = "Minor Main"
        ElseIf oInvoice.FeeType = "2" Then
            lblFeeTypeValue.Text = "Late Fee"
        ElseIf oInvoice.FeeType = "3" Then
            lblFeeTypeValue.Text = "Miscellaneous"
        ElseIf oInvoice.FeeType = "C" Then
            lblFeeTypeValue.Text = "Credit Memo"
        Else
            lblFeeTypeValue.Text = oInvoice.FeeType
        End If

        LoadugGenCreditMemo()
        If bolUpdateCM Then
            Me.Text = "Update Credit Memo"
            btnIssueCreditMemo.Text = "Update Credit Memo"
            btnDeleteCreditMemo.Visible = True
        Else
            btnDeleteCreditMemo.Visible = True
        End If

    End Sub
#End Region
#Region " UI Support Routines"
    'Private Sub frmClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
    '    If sender.GetType.Name.IndexOf("CreditMemoSummary") >= 0 Then
    '    End If
    'End Sub
    'Private Sub frmClosed(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If sender.GetType.Name.IndexOf("CreditMemoSummary") >= 0 Then
    '        frmCreditMemoSummary = Nothing
    '    End If
    'End Sub
    Private Sub LoadugGenCreditMemo()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16


        dsLocal = oFeeInvoice.GetInvoiceLineItemSummaryGrid_ByInvoiceID(InvoiceID, bolUpdateCM)
        dsLocal.Tables(0).Columns("Facility_ID").ReadOnly = True
        dsLocal.Tables(0).Columns("FacilityName").ReadOnly = True
        dsLocal.Tables(0).Columns("Fiscal_Year").ReadOnly = True
        dsLocal.Tables(0).Columns("Charges").ReadOnly = True
        dsLocal.Tables(0).Columns("Balance").ReadOnly = True

        ugGenCreditMemo.DataSource = dsLocal

        'uggencreditmemo

        ugGenCreditMemo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows

            ugGenCreditMemo.DisplayLayout.Bands(0).SortedColumns.Clear()
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("Fiscal_Year").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("Quantity").Hidden = True
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("UnitPrice").Hidden = True
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("ITEM_SEQ_NUMBER").Hidden = True
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("CM_INV__ID").Hidden = True
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("CM_LINE_INV__ID").Hidden = True

            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("Facility_ID").TabStop = False
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("FacilityName").TabStop = False
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("Fiscal_Year").TabStop = False
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("Charges").TabStop = False
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("Balance").TabStop = False
            ugGenCreditMemo.DisplayLayout.Bands(0).Columns("LineDescription").FieldLen = 50
        End If

        Dim sTotal As Single = 0.0
        For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugGenCreditMemo.Rows
            If IsNumeric(ugRow.Cells("Credits").Value) Then
                If ugRow.Cells("Credits").Value > 0 Then
                    sTotal = sTotal + ugRow.Cells("Credits").Value
                End If
            Else
                ugRow.Cells("Credits").Value = 0
            End If
        Next
        Me.lblCreditAmtValue.Text = FormatNumber(sTotal, 2, TriState.True, TriState.False, TriState.True)

    End Sub
    Private Sub ProcessCreditMemo()
        Dim oInvoiceHeader As New MUSTER.BusinessLogic.pFeeInvoice
        Dim oInvoiceLineItem As New MUSTER.Info.FeeInvoiceInfo
        Dim iSequence As Int16
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        Try
            If Not ValidateData() Then Exit Sub

            If bolUpdateCM And Not ugGenCreditMemo.Rows(0).Cells("CM_INV__ID").Value Is DBNull.Value Then
                oInvoiceHeader.Retrieve(ugGenCreditMemo.Rows(0).Cells("CM_INV__ID").Value)
                oInvoiceHeader.InvoiceLineItems = oInvoiceHeader.RetrieveLineItems(oInvoiceHeader.ID)
            Else
                oInvoiceHeader.InvoiceType = "C"
                oInvoiceHeader.FeeType = "C"
                oInvoiceHeader.RecType = "ADVIC"
                oInvoiceHeader.OwnerID = OwnerID
                oInvoiceHeader.CreditApplyTo = oInvoice.WarrantNumber
                oInvoiceHeader.Description = "Credit Memo For " & oInvoice.WarrantNumber
                'Get Fiscal Year from Fees Basis
                Dim oFeeBasis As New MUSTER.BusinessLogic.pFeeBasis
                oInvoiceHeader.FiscalYear = oFeeBasis.GetFiscalYearForFee(Now.Date)
            End If
            oInvoiceHeader.InvoiceAmount = lblCreditAmtValue.Text
            iSequence = 0
            If ugGenCreditMemo.Rows.Count > 0 Then
                For Each ugrow In ugGenCreditMemo.Rows
                    If ugrow.Cells("CM_LINE_INV__ID").Value Is DBNull.Value Then
                        If ugrow.Cells("Credits").Value > 0 Then
                            iSequence += 1
                            oInvoiceLineItem = New MUSTER.Info.FeeInvoiceInfo
                            oInvoiceLineItem.InvoiceType = "C"
                            oInvoiceLineItem.FeeType = "C"
                            oInvoiceLineItem.RecType = "ADLN"
                            oInvoiceLineItem.ID = iSequence * -1
                            'Issue #2755 - credit memo sfy should be same as sfy for the line item the credit memo is being issued
                            oInvoiceLineItem.FiscalYear = ugrow.Cells("Fiscal_Year").Value
                            oInvoiceLineItem.OwnerID = OwnerID
                            oInvoiceLineItem.FacilityID = ugrow.Cells("Facility_ID").Value
                            oInvoiceLineItem.SequenceNumber = ugrow.Cells("ITEM_SEQ_NUMBER").Value
                            If Not IsDBNull(ugrow.Cells("LineDescription").Value) Then
                                If ugrow.Cells("LineDescription").Value = "" Then
                                    oInvoiceLineItem.Description = "Credit Memo For " & oInvoice.WarrantNumber & " FacID: " & ugrow.Cells("Facility_ID").Value
                                Else
                                    If ugrow.Cells("LineDescription").Text.Length > 50 Then
                                        oInvoiceLineItem.Description = ugrow.Cells("LineDescription").Text.Substring(0, 50)
                                    Else
                                        oInvoiceLineItem.Description = ugrow.Cells("LineDescription").Text
                                    End If
                                End If
                            Else
                                oInvoiceLineItem.Description = "Credit Memo For " & oInvoice.WarrantNumber & " FacID: " & ugrow.Cells("Facility_ID").Value
                            End If
                            oInvoiceLineItem.InvoiceLineAmount = ugrow.Cells("Credits").Value
                            oInvoiceLineItem.Quantity = ugrow.Cells("Quantity").Value
                            oInvoiceLineItem.UnitPrice = ugrow.Cells("UnitPrice").Value
                            oInvoiceHeader.InvoiceLineItems.Add(oInvoiceLineItem)
                        End If
                    Else
                        oInvoiceLineItem = oInvoiceHeader.InvoiceLineItems.Item(ugrow.Cells("CM_LINE_INV__ID").Value)
                        If oInvoiceLineItem Is Nothing Then
                            Dim oInvoiceLocal As New MUSTER.BusinessLogic.pFeeInvoice
                            oInvoiceLineItem = oInvoiceLocal.Retrieve(ugrow.Cells("CM_LINE_INV__ID").Value)
                            If oInvoiceLineItem Is Nothing Then
                                MsgBox("Credit Memo (LineItem) not found for FacID: " + ugrow.Cells("Facility_ID").Value.ToString, MsgBoxStyle.OKOnly, "Line Item Not Found")
                                Exit Sub
                            ElseIf oInvoiceLineItem.ID <> ugrow.Cells("CM_LINE_INV__ID").Value Then
                                MsgBox("Credit Memo (LineItem) not found for FacID: " + ugrow.Cells("Facility_ID").Value.ToString, MsgBoxStyle.OKOnly, "Line Item Not Found")
                                Exit Sub
                            End If
                            oInvoiceHeader.InvoiceLineItems.Add(oInvoiceLineItem)
                        End If
                        If ugrow.Cells("Credits").Value > 0 Then
                            oInvoiceLineItem.InvoiceLineAmount = ugrow.Cells("Credits").Value
                            If Not IsDBNull(ugrow.Cells("LineDescription").Value) Then
                                If ugrow.Cells("LineDescription").Value = "" Then
                                    oInvoiceLineItem.Description = "Credit Memo For " & oInvoice.WarrantNumber & " FacID: " & ugrow.Cells("Facility_ID").Value
                                Else
                                    If ugrow.Cells("LineDescription").Text.Length > 50 Then
                                        oInvoiceLineItem.Description = ugrow.Cells("LineDescription").Text.Substring(0, 50)
                                    Else
                                        oInvoiceLineItem.Description = ugrow.Cells("LineDescription").Text
                                    End If
                                End If
                            Else
                                oInvoiceLineItem.Description = "Credit Memo For " & oInvoice.WarrantNumber & " FacID: " & ugrow.Cells("Facility_ID").Value
                            End If
                        Else
                            oInvoiceLineItem.Deleted = True
                        End If
                    End If
                Next
            End If

            If oInvoiceHeader.ID <= 0 Then
                oInvoiceHeader.CreatedBy = MusterContainer.AppUser.ID
            Else
                oInvoiceHeader.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oInvoiceHeader.SaveNewInvoice(UIUtilsGen.ModuleID.Fees, MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            MsgBox("Credit Memo Saved Successfully", MsgBoxStyle.OKOnly, "Save Success")
            If Not CallingForm Is Nothing Then CallingForm.Tag = "1"
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub DeleteCreditMemo()
        Dim oInvoiceHeader As New MUSTER.BusinessLogic.pFeeInvoice
        Dim oInvoiceLineItem As New MUSTER.Info.FeeInvoiceInfo
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        If ugGenCreditMemo.Rows(0).Cells("CM_INV__ID").Value Is DBNull.Value Then
            MsgBox("Invalid Credit Memo Invoice ID. Please contact Administrator", MsgBoxStyle.OKOnly, "Invalid Invoice ID")
            Exit Sub
        End If
        oInvoiceHeader.Retrieve(ugGenCreditMemo.Rows(0).Cells("CM_INV__ID").Value)
        oInvoiceHeader.Deleted = True
        oInvoiceHeader.ModifiedBy = MusterContainer.AppUser.ID
        oInvoiceHeader.Save(UIUtilsGen.ModuleID.Fees, MusterContainer.AppUser.UserKey, returnVal, "FEES")
        If Not UIUtilsGen.HasRights(returnVal) Then
            Exit Sub
        End If
        MsgBox("Deleted Credit Memo Successfully", MsgBoxStyle.OKOnly, "Delete Success")
        If Not CallingForm Is Nothing Then CallingForm.Tag = "1"
        Me.Close()
    End Sub
    Private Function ValidateData() As Boolean
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim bolReturn As Boolean
        bolReturn = True
        Try

            If ugGenCreditMemo.Rows.Count = 0 Then
                MsgBox("Invoice has no charges to credit.", MsgBoxStyle.OKOnly, "Invalid Credit Amount")
                Return False
            End If

            If lblCreditAmtValue.Text = "0.00" Then
                MsgBox("Credit Memo Has No Credit Amount", MsgBoxStyle.OKOnly, "Invalid Credit Amount")
                Return False
            End If
            For Each ugrow In ugGenCreditMemo.Rows
                If ugrow.Cells("Credits").Value > 0 Then
                    If ugrow.Cells("Credits").Value > ugrow.Cells("Charges").Value Then
                        MsgBox("Credit cannot be larger than the charge for Sequence Number (" & ugrow.Cells("ITEM_SEQ_NUMBER").Value & ").", MsgBoxStyle.OKOnly, "Invalid Credit Amount")
                        Return False
                    End If
                    If Not IsDBNull(ugrow.Cells("LineDescription").Value) Then
                        If ugrow.Cells("LineDescription").Value = "" Then
                            MsgBox("Line Description Required for Sequence Number (" & ugrow.Cells("ITEM_SEQ_NUMBER").Value & ").", MsgBoxStyle.OKOnly, "Invalid Reason")
                            Return False
                        End If
                    Else
                        MsgBox("Line Description Required for Sequence Number (" & ugrow.Cells("ITEM_SEQ_NUMBER").Value & ").", MsgBoxStyle.OKOnly, "Invalid Reason")
                        Return False
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Function
#End Region
#Region " UI Control Events"
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
    Private Sub btnIssueCreditMemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIssueCreditMemo.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim balanceFlag As Short
        balanceFlag = 0 'if 0, means have 0 balance
        For Each ugrow In ugGenCreditMemo.Rows
            If ugrow.Cells("Balance").Value > 0 Then
                balanceFlag = 1
            End If
        Next
        If balanceFlag = 0 Then
            MsgBox("Cannot issue credit for an invoice with 0.00 balance.")
            btnIssueCreditMemo.Enabled = False
            Exit Sub
        Else
            btnIssueCreditMemo.Enabled = True
        End If
        ProcessCreditMemo()
    End Sub
    Private Sub btnDeleteCreditMemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteCreditMemo.Click
        DeleteCreditMemo()
    End Sub
    Private Sub ugGenCreditMemo_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugGenCreditMemo.AfterCellUpdate
        Dim sTotal As Single = 0.0
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        For Each ugrow In ugGenCreditMemo.Rows
            If IsNumeric(ugrow.Cells("Credits").Value) Then
                If ugrow.Cells("Credits").Value > 0 Then
                    'ugrow.Cells("Credits").Value = FormatNumber(ugrow.Cells("Credits").Value, 2, TriState.True, TriState.False, TriState.True)
                    sTotal = sTotal + ugrow.Cells("Credits").Value
                End If
            Else
                ugrow.Cells("Credits").Value = 0
            End If
            If IsDBNull(ugrow.Cells("LineDescription").Value) = False Then
                If Len(ugrow.Cells("LineDescription").Value) > 50 Then
                    MsgBox("Line Description Exceeds 50 characters for Sequence Number (" & ugrow.Cells("ITEM_SEQ_NUMBER").Value & ").  Line will be truncated.", MsgBoxStyle.OKOnly, "Invalid Line Length")
                End If
            End If
        Next
        Me.lblCreditAmtValue.Text = FormatNumber(sTotal, 2, TriState.True, TriState.False, TriState.True)
    End Sub
#End Region
End Class
