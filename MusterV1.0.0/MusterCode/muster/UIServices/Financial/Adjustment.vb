Public Class Adjustment
    Inherits System.Windows.Forms.Form

    Friend FinancialEventID As Int64
    Friend FinancialCommitmentID As Int64
    Friend AdjustmentID As Int64
    Friend Balance As Double
    Friend NegativeRequest As Boolean = False

    Private oFinancialEvent As New MUSTER.BusinessLogic.pFinancial
    Private oFinancialCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
    Private oFinancialAdjustment As New MUSTER.BusinessLogic.pFinancialCommitAdjustment
    Private nContactID As Integer = 0


    Private bolFormatting As Boolean
    Private bolLoading As Boolean
    Private sCommitmentTotal As Double
    Private sStartingTotal As Double
    Dim returnVal As String = String.Empty
    Friend nAdjustmentID As Integer = 0
    Friend ugCommitRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Friend SystemComment As String = String.Empty
    Dim strPrevType As String = String.Empty

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal ContactID As Integer)
        MyBase.New()

        Me.nContactID = ContactID

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
    Friend WithEvents pnlAdjustmentBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlAdjustment As System.Windows.Forms.Panel
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents lblIDValue As System.Windows.Forms.Label
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents dtPickDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents lblAmount As System.Windows.Forms.Label
    Friend WithEvents chkDirectorAppReqd As System.Windows.Forms.CheckBox
    Friend WithEvents chkFinancialAppReqd As System.Windows.Forms.CheckBox
    Friend WithEvents chkApproved As System.Windows.Forms.CheckBox
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents lblCommit As System.Windows.Forms.Label
    Friend WithEvents lblCommitment As System.Windows.Forms.Label
    Friend WithEvents cfCostFormat As MUSTER.CostFormat
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlAdjustmentBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.pnlAdjustment = New System.Windows.Forms.Panel
        Me.cfCostFormat = New MUSTER.CostFormat
        Me.lblCommitment = New System.Windows.Forms.Label
        Me.lblComments = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.chkApproved = New System.Windows.Forms.CheckBox
        Me.chkFinancialAppReqd = New System.Windows.Forms.CheckBox
        Me.chkDirectorAppReqd = New System.Windows.Forms.CheckBox
        Me.lblAmount = New System.Windows.Forms.Label
        Me.txtAmount = New System.Windows.Forms.TextBox
        Me.lblType = New System.Windows.Forms.Label
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.dtPickDate = New System.Windows.Forms.DateTimePicker
        Me.lblDate = New System.Windows.Forms.Label
        Me.lblCommit = New System.Windows.Forms.Label
        Me.lblIDValue = New System.Windows.Forms.Label
        Me.lblID = New System.Windows.Forms.Label
        Me.pnlAdjustmentBottom.SuspendLayout()
        Me.pnlAdjustment.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlAdjustmentBottom
        '
        Me.pnlAdjustmentBottom.Controls.Add(Me.btnCancel)
        Me.pnlAdjustmentBottom.Controls.Add(Me.btnOK)
        Me.pnlAdjustmentBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlAdjustmentBottom.Location = New System.Drawing.Point(0, 310)
        Me.pnlAdjustmentBottom.Name = "pnlAdjustmentBottom"
        Me.pnlAdjustmentBottom.Size = New System.Drawing.Size(376, 40)
        Me.pnlAdjustmentBottom.TabIndex = 8
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(191, 9)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(111, 9)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "OK"
        '
        'pnlAdjustment
        '
        Me.pnlAdjustment.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlAdjustment.Controls.Add(Me.cfCostFormat)
        Me.pnlAdjustment.Controls.Add(Me.lblCommitment)
        Me.pnlAdjustment.Controls.Add(Me.lblComments)
        Me.pnlAdjustment.Controls.Add(Me.txtComments)
        Me.pnlAdjustment.Controls.Add(Me.chkApproved)
        Me.pnlAdjustment.Controls.Add(Me.chkFinancialAppReqd)
        Me.pnlAdjustment.Controls.Add(Me.chkDirectorAppReqd)
        Me.pnlAdjustment.Controls.Add(Me.lblAmount)
        Me.pnlAdjustment.Controls.Add(Me.txtAmount)
        Me.pnlAdjustment.Controls.Add(Me.lblType)
        Me.pnlAdjustment.Controls.Add(Me.cmbType)
        Me.pnlAdjustment.Controls.Add(Me.dtPickDate)
        Me.pnlAdjustment.Controls.Add(Me.lblDate)
        Me.pnlAdjustment.Controls.Add(Me.lblCommit)
        Me.pnlAdjustment.Controls.Add(Me.lblIDValue)
        Me.pnlAdjustment.Controls.Add(Me.lblID)
        Me.pnlAdjustment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlAdjustment.Location = New System.Drawing.Point(0, 0)
        Me.pnlAdjustment.Name = "pnlAdjustment"
        Me.pnlAdjustment.Size = New System.Drawing.Size(376, 310)
        Me.pnlAdjustment.TabIndex = 1
        '
        'cfCostFormat
        '
        Me.cfCostFormat.Location = New System.Drawing.Point(328, 48)
        Me.cfCostFormat.Name = "cfCostFormat"
        Me.cfCostFormat.Size = New System.Drawing.Size(600, 232)
        Me.cfCostFormat.TabIndex = 16
        Me.cfCostFormat.Visible = False
        '
        'lblCommitment
        '
        Me.lblCommitment.BackColor = System.Drawing.SystemColors.Window
        Me.lblCommitment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCommitment.Location = New System.Drawing.Point(80, 32)
        Me.lblCommitment.Name = "lblCommitment"
        Me.lblCommitment.Size = New System.Drawing.Size(272, 23)
        Me.lblCommitment.TabIndex = 15
        Me.lblCommitment.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblComments
        '
        Me.lblComments.Location = New System.Drawing.Point(16, 216)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(62, 17)
        Me.lblComments.TabIndex = 14
        Me.lblComments.Text = "Comments:"
        Me.lblComments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(80, 216)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtComments.Size = New System.Drawing.Size(280, 64)
        Me.txtComments.TabIndex = 7
        Me.txtComments.Text = ""
        '
        'chkApproved
        '
        Me.chkApproved.Location = New System.Drawing.Point(64, 184)
        Me.chkApproved.Name = "chkApproved"
        Me.chkApproved.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkApproved.TabIndex = 6
        Me.chkApproved.Text = "Approved"
        '
        'chkFinancialAppReqd
        '
        Me.chkFinancialAppReqd.Location = New System.Drawing.Point(0, 160)
        Me.chkFinancialAppReqd.Name = "chkFinancialAppReqd"
        Me.chkFinancialAppReqd.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkFinancialAppReqd.Size = New System.Drawing.Size(168, 24)
        Me.chkFinancialAppReqd.TabIndex = 5
        Me.chkFinancialAppReqd.Text = "Financial Approval Required"
        '
        'chkDirectorAppReqd
        '
        Me.chkDirectorAppReqd.Enabled = False
        Me.chkDirectorAppReqd.Location = New System.Drawing.Point(8, 136)
        Me.chkDirectorAppReqd.Name = "chkDirectorAppReqd"
        Me.chkDirectorAppReqd.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkDirectorAppReqd.Size = New System.Drawing.Size(160, 24)
        Me.chkDirectorAppReqd.TabIndex = 4
        Me.chkDirectorAppReqd.Text = "Director Approval Required"
        '
        'lblAmount
        '
        Me.lblAmount.Location = New System.Drawing.Point(32, 112)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.Size = New System.Drawing.Size(46, 23)
        Me.lblAmount.TabIndex = 9
        Me.lblAmount.Text = "Amount:"
        Me.lblAmount.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(80, 112)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.TabIndex = 3
        Me.txtAmount.Text = ""
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(48, 88)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(32, 23)
        Me.lblType.TabIndex = 7
        Me.lblType.Text = "Type:"
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbType
        '
        Me.cmbType.Location = New System.Drawing.Point(80, 88)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(136, 21)
        Me.cmbType.TabIndex = 2
        '
        'dtPickDate
        '
        Me.dtPickDate.Checked = False
        Me.dtPickDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickDate.Location = New System.Drawing.Point(80, 64)
        Me.dtPickDate.Name = "dtPickDate"
        Me.dtPickDate.Size = New System.Drawing.Size(88, 20)
        Me.dtPickDate.TabIndex = 1
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(48, 64)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(32, 17)
        Me.lblDate.TabIndex = 4
        Me.lblDate.Text = "Date:"
        Me.lblDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCommit
        '
        Me.lblCommit.Location = New System.Drawing.Point(8, 32)
        Me.lblCommit.Name = "lblCommit"
        Me.lblCommit.Size = New System.Drawing.Size(72, 17)
        Me.lblCommit.TabIndex = 2
        Me.lblCommit.Text = "Commitment:"
        Me.lblCommit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblIDValue
        '
        Me.lblIDValue.BackColor = System.Drawing.SystemColors.Window
        Me.lblIDValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblIDValue.Location = New System.Drawing.Point(80, 8)
        Me.lblIDValue.Name = "lblIDValue"
        Me.lblIDValue.TabIndex = 1
        Me.lblIDValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblID
        '
        Me.lblID.Location = New System.Drawing.Point(56, 9)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(22, 17)
        Me.lblID.TabIndex = 0
        Me.lblID.Text = "ID:"
        Me.lblID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Adjustment
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 350)
        Me.Controls.Add(Me.pnlAdjustment)
        Me.Controls.Add(Me.pnlAdjustmentBottom)
        Me.MaximizeBox = False
        Me.Name = "Adjustment"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Adjustment"
        Me.pnlAdjustmentBottom.ResumeLayout(False)
        Me.pnlAdjustment.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub txtAmount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAmount.TextChanged
        If bolLoading = True Or bolFormatting = True Then Exit Sub
        If Not IsNumeric(txtAmount.Text) Then
            txtAmount.Text = 0
        End If
        oFinancialAdjustment.AdjustAmount = CDbl(txtAmount.Text)
    End Sub
    Private Sub cmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        If bolLoading = True Then Exit Sub
        If cmbType.Text = "Change Order" Then
            chkFinancialAppReqd.Enabled = True
            chkApproved.Enabled = True
            If Not (oFinancialAdjustment.CommitAdjustmentID = 0 And oFinancialCommitment.ActivityType = 19) Then
                If CDbl(txtAmount.Text) > CDbl(sCommitmentTotal * 0.25) Then
                    chkDirectorAppReqd.Checked = True
                Else
                    chkDirectorAppReqd.Checked = False
                End If
            End If
        Else
            chkFinancialAppReqd.Checked = False
            chkFinancialAppReqd.Enabled = False
            chkApproved.Enabled = False
            chkApproved.Checked = False
            chkDirectorAppReqd.Checked = False
        End If
        oFinancialAdjustment.AdjustType = cmbType.SelectedValue

    End Sub

    Private Sub dtPickDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickDate.ValueChanged
        If bolLoading = True Then Exit Sub
        oFinancialAdjustment.AdjustDate = dtPickDate.Value.Date

    End Sub

    Private Sub chkDirectorAppReqd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDirectorAppReqd.CheckedChanged
        If bolLoading = True Then Exit Sub
        oFinancialAdjustment.DirectorApprovalReq = chkDirectorAppReqd.Checked

    End Sub

    Private Sub chkFinancialAppReqd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFinancialAppReqd.CheckedChanged
        If bolLoading = True Then Exit Sub
        oFinancialAdjustment.FinancialApprovalReq = chkFinancialAppReqd.Checked

    End Sub

    Private Sub chkApproved_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkApproved.CheckedChanged
        If bolLoading = True Then Exit Sub
        oFinancialAdjustment.Approved = chkApproved.Checked

    End Sub

    Private Sub txtComments_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtComments.TextChanged
        If bolLoading = True Then Exit Sub
        oFinancialAdjustment.Comments = txtComments.Text
    End Sub

    Private Sub txtAmount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmount.LostFocus
        If bolLoading = True Or bolFormatting = True Then Exit Sub
        If CDbl(txtAmount.Text) <= 0.0 Then
            MsgBox("Adjustment Amount Must Be Greater Than Zero.")
        End If

        If cmbType.Text = "Change Order" Then
            If Not (oFinancialAdjustment.CommitAdjustmentID = 0 And oFinancialCommitment.ActivityType = 19) Then
                If CDbl(txtAmount.Text) > CDbl(sCommitmentTotal * 0.25) Then
                    chkDirectorAppReqd.Checked = True
                Else
                    chkDirectorAppReqd.Checked = False
                End If
            End If
        End If

        bolFormatting = True
        txtAmount.Text = FormatNumber(txtAmount.Text, 2, TriState.False, TriState.False, TriState.True)
        bolFormatting = False
    End Sub

    Private Sub Adjustment_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oFinancialEvent.Retrieve(FinancialEventID)
        oFinancialCommitment.Retrieve(FinancialCommitmentID)
        oFinancialAdjustment.Retrieve(AdjustmentID)


        cfCostFormat.AssignCommitmentObject(oFinancialCommitment)
        cfCostFormat.CostFormatType = oFinancialCommitment.Case_Letter
        cfCostFormat.SetDisplay(False)
        cfCostFormat.LoadCommitment()
        sCommitmentTotal = cfCostFormat.GrandTotal
        bolLoading = True
        LoadDropDowns()
        FormSetup()
        bolLoading = False

        If Not SystemComment = String.Empty Then
            txtComments.Text = SystemComment
        End If

    End Sub



    Private Sub FormSetup()
        Dim oFinancialActivity As New MUSTER.BusinessLogic.pFinancialActivity

        oFinancialActivity.Retrieve(oFinancialCommitment.ActivityType)
        If oFinancialAdjustment.CommitAdjustmentID = 0 Then
            lblIDValue.Text = "New"
            txtAmount.Text = "0.00"
            chkDirectorAppReqd.Checked = False
            chkFinancialAppReqd.Checked = False
            chkApproved.Checked = False
            dtPickDate.Value = Now.Date
            oFinancialAdjustment.AdjustAmount = 0.0
            oFinancialAdjustment.AdjustType = cmbType.SelectedValue
            oFinancialAdjustment.AdjustDate = Now.Date
            oFinancialAdjustment.Approved = False
            oFinancialAdjustment.Comments = ""
            oFinancialAdjustment.CommitmentID = FinancialCommitmentID
            oFinancialAdjustment.DirectorApprovalReq = False
            oFinancialAdjustment.FinancialApprovalReq = False

            If Balance <> 0 Then
                If Balance > 0 Then
                    cmbType.SelectedValue = 1076
                    oFinancialAdjustment.AdjustAmount = Balance
                    oFinancialAdjustment.AdjustType = cmbType.SelectedValue
                Else
                    If NegativeRequest Then
                        cmbType.SelectedValue = 1075
                        oFinancialAdjustment.AdjustAmount = Balance
                        oFinancialAdjustment.AdjustType = cmbType.SelectedValue
                    Else
                        cmbType.SelectedValue = 1075
                        oFinancialAdjustment.AdjustAmount = Balance * -1
                        oFinancialAdjustment.AdjustType = cmbType.SelectedValue
                    End If
                End If
                txtAmount.Text = FormatNumber(oFinancialAdjustment.AdjustAmount, 2, TriState.True, TriState.False, TriState.True)
            End If
        Else
            lblIDValue.Text = oFinancialAdjustment.CommitAdjustmentID
            txtAmount.Text = FormatNumber(oFinancialAdjustment.AdjustAmount, 2, TriState.True, TriState.False, TriState.True)
            cmbType.SelectedValue = oFinancialAdjustment.AdjustType
            dtPickDate.Value = oFinancialAdjustment.AdjustDate
            chkApproved.Checked = oFinancialAdjustment.Approved
            chkDirectorAppReqd.Checked = oFinancialAdjustment.DirectorApprovalReq
            chkFinancialAppReqd.Checked = oFinancialAdjustment.FinancialApprovalReq
            txtComments.Text = oFinancialAdjustment.Comments

            If cmbType.Text = "Change Order" Then
                chkFinancialAppReqd.Enabled = True
                chkApproved.Enabled = True

            Else
                chkFinancialAppReqd.Checked = False
                chkFinancialAppReqd.Enabled = False
                chkApproved.Enabled = False
                chkApproved.Checked = False
            End If
            strPrevType = cmbType.Text
        End If

        sStartingTotal = txtAmount.Text - Balance
        lblCommitment.Text = oFinancialCommitment.ApprovedDate.ToShortDateString & " - " & oFinancialCommitment.PONumber & " - " & oFinancialActivity.ActivityDescShort & " - $" & FormatNumber(sCommitmentTotal, 2, TriState.True, TriState.False, TriState.True)

    End Sub

    Private Sub LoadDropDowns()

        Try
            Dim dtAdjustmentTypes As DataTable = oFinancialAdjustment.PopulateFinancialCommitmentAdjustmentTypes


            cmbType.DataSource = dtAdjustmentTypes
            cmbType.DisplayMember = "PROPERTY_NAME"
            cmbType.ValueMember = "PROPERTY_ID"


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try


    End Sub
    Private Function ValidateNegativeBalance() As Boolean
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        Dim sTotalAdjustAmount As Double = 0.0
        Dim sTotalChangeOrderAmt As Double = 0.0
        Dim sTotalunencumberanceAmt As Double = 0.0
        Dim sTotalPaid As Double = 0.0
        If ugCommitRow Is Nothing Then
            Return True
            Exit Function
        End If
        Try
            If Not ugCommitRow.ChildBands Is Nothing Then
                For Each ChildBand In ugCommitRow.ChildBands
                    If Not ChildBand.Rows Is Nothing Then
                        If ChildBand.Rows.Count > 0 Then
                            For Each Childrow In ChildBand.Rows  ' Adjustment
                                If Childrow.Cells("CommitmentID").Value = ugCommitRow.Cells("CommitmentID").Value Then
                                    If Childrow.Cells.Exists("Adjust_Type") Then
                                        If Not Childrow.Cells("ChildID").Value = oFinancialAdjustment.CommitAdjustmentID Then
                                            If UCase(Childrow.Cells("Adjust_Type").Value) = UCase("change order") Then
                                                If Not Childrow.Cells("Adjust_Amount").Value Is System.DBNull.Value Then
                                                    sTotalChangeOrderAmt += CDbl(Childrow.Cells("Adjust_Amount").Value)
                                                End If
                                            ElseIf UCase(Childrow.Cells("Adjust_Type").Value) = UCase("unencumberance") Then
                                                If Not Childrow.Cells("Adjust_Amount").Value Is System.DBNull.Value Then
                                                    sTotalunencumberanceAmt += CDbl(Childrow.Cells("Adjust_Amount").Value)
                                                End If
                                            End If
                                        End If
                                    Else   ' Invoice Payment
                                        If Not Childrow.Cells("Paid").Value Is System.DBNull.Value Then
                                            sTotalPaid += CDbl(Childrow.Cells("Paid").Value)
                                        End If
                                    End If

                                End If
                            Next
                        End If
                    End If
                Next
            End If

            If cmbType.Text <> "Change Order" Then
                sTotalunencumberanceAmt += oFinancialAdjustment.AdjustAmount
            Else
                sTotalChangeOrderAmt += oFinancialAdjustment.AdjustAmount
            End If
            sTotalAdjustAmount = sTotalChangeOrderAmt - sTotalunencumberanceAmt
            'Balance = Commitment + Adjustment - Payment
            Dim sBalance As Double
            sBalance = IIf(ugCommitRow.Cells("Commitment").Value Is DBNull.Value, 0.0, CDbl(ugCommitRow.Cells("Commitment").Value)) + sTotalAdjustAmount - sTotalPaid
            sBalance = FormatNumber(sBalance, 2)
            If sBalance < 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Balance Validation Error " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Function
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim bolApprovedStatus As Boolean = False
        bolApprovedStatus = oFinancialAdjustment.ApprovedOriginal
        If CDbl(txtAmount.Text) <= 0.0 Then
            MsgBox("Adjustment Amount Must Be Greater Than Zero.")
            Exit Sub
        End If
        If cmbType.Text = "Change Order" Then
            If txtComments.Text.Trim = "" Then
                MsgBox("Adjustment Reason/Comment Required.")
                Exit Sub
            End If
        End If

        'If GetCommitmentTotals(True) < 0 Then
        '    MsgBox("This adjustment would result in a negative balance.  Save not allowed.")
        '    Exit Sub
        'End If
        Dim bolCheck800000 As Boolean = True
        ' #2925
        If GetCommitmentTotals(False) > 1500000 Then
            bolCheck800000 = False
            If MusterContainer.AppUser.HEAD_FINANCIAL And cmbType.Text.IndexOf("unencumberance") > -1 Then
                If Not MsgBox("This adjustment would result in a commitment greater than 1,500,000. Do you want to override?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Exit Sub
                End If
            Else
                MsgBox("This adjustment would result in a commitment greater than 1,500,000. Save not allowed.")
                Exit Sub
            End If
        End If
        If bolCheck800000 Then
            If GetCommitmentTotals(False) > 1200000 Then
                MsgBox("This adjustment will result in a commitment greater than 1,200,000. Save is allowed.", MsgBoxStyle.Exclamation)
            End If
        End If

        If Not ValidateNegativeBalance() Then
            MsgBox("Balance Cannot Be Less Than Zero.")
            Exit Sub
        End If

        If cmbType.Text = "Change Order" Then

            If chkDirectorAppReqd.Checked = False Then
                If oFinancialCommitment.ActivityType <> 19 Then
                    If CDbl(txtAmount.Text) > CDbl(sCommitmentTotal * 0.25) Then
                        chkDirectorAppReqd.Checked = True
                    End If
                End If
            End If
        End If

        If Balance < 0 Then
            If oFinancialAdjustment.AdjustAmount < (Balance * -1) Then
                MsgBox("Adjustment Amount Must Be Greater Than or Equal to:  " & CStr((Balance * -1)) & " ", MsgBoxStyle.OKOnly, "Invalid Amount")
                Exit Sub
            End If
        End If
        If Balance > 0 Then
            If oFinancialAdjustment.AdjustAmount > Balance Then
                MsgBox("Adjustment Amount Must Be Less Than or Equal to:  " & CStr(Balance) & " ", MsgBoxStyle.OKOnly, "Invalid Amount")
                Exit Sub
            End If
        End If

        If Not SystemComment = String.Empty And txtComments.Text = String.Empty Then
            txtComments.Text = SystemComment
        End If


        If oFinancialAdjustment.CommitAdjustmentID <= 0 Then
            oFinancialAdjustment.CreatedBy = MusterContainer.AppUser.ID
        Else
            oFinancialAdjustment.ModifiedBy = MusterContainer.AppUser.ID
        End If
        oFinancialAdjustment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
        nAdjustmentID = oFinancialAdjustment.CommitAdjustmentID
        If Not UIUtilsGen.HasRights(returnVal) Then
            Exit Sub
        End If

        If cmbType.Text = "Change Order" Then

            If lblIDValue.Text = "New" Then
                GenerateLetter("ChangeOrderMemoTemplate.doc", "Change Order Encumberance Memo", "EncumberanceMemo", nContactID)
            Else
                If Not bolApprovedStatus Then
                    If MsgBox("Do you want to regenerate the Change Order Encumberance Memo?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        GenerateLetter("ChangeOrderMemoTemplate.doc", "Change Order Encumberance Memo", "EncumberanceMemo", nContactID)
                    End If
                End If
            End If
        Else

            If lblIDValue.Text = "New" Then
                GenerateLetter("UnencumberanceMemoTemplate.doc", "Change Order Unencumberance Memo", "UnencumberanceMemo", nContactID)
            Else
                If MsgBox("Do you want to regenerate the Change Order Unencumberance Memo?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    GenerateLetter("UnencumberanceMemoTemplate.doc", "Change Order Unencumberance Memo", "UnencumberanceMemo", nContactID)
                End If
            End If
        End If

        If chkDirectorAppReqd.Checked = True Then

            If lblIDValue.Text = "New" Then
                GenerateLetter("ChangeOrderForm.doc", "Executive Director Change Order Approval Form", "ChangeOrderApproval", nContactID)

            Else
                If Not bolApprovedStatus Then
                    If MsgBox("Do you want to regenerate the Executive Director Change Order Approval Form?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        GenerateLetter("ChangeOrderForm.doc", "Executive Director Change Order Approval Form", "ChangeOrderApproval", nContactID)
                    End If
                End If
            End If
        End If
        Me.Close()

    End Sub

    Private Sub GenerateLetter(ByVal strTemplate As String, ByVal strLongName As String, ByVal strShortName As String, ByVal contactID As Integer)
        Dim oLetter As New Reg_Letters
        Dim oCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
        Dim oFinancial As New MUSTER.BusinessLogic.pFinancial
        Dim oTechEvent As New MUSTER.BusinessLogic.pLustEvent
        Dim oFacility As New MUSTER.BusinessLogic.pFacility
        Dim oOwner As New MUSTER.BusinessLogic.pOwner

        oCommitment.Retrieve(oFinancialAdjustment.CommitmentID)
        oFinancial.Retrieve(oCommitment.Fin_Event_ID)
        oTechEvent.Retrieve(oFinancial.TecEventID)
        oFacility.Retrieve(oOwner.OwnerInfo, oTechEvent.FacilityID, "SELF", "FACILITY")
        oOwner.Retrieve(oFacility.OwnerID)

        oLetter.GenerateFinancialLetter(oTechEvent.FacilityID, strLongName, strShortName, strLongName, strTemplate, oOwner, oFinancial.TecEventID, oCommitment.Fin_Event_ID, oFinancialAdjustment.CommitmentID, 0, oFinancialAdjustment.CommitAdjustmentID, 0, , oFinancial.ID, oFinancial.Sequence, UIUtilsGen.EntityTypes.FinancialEvent, contactid)

    End Sub

    Private Function GetCommitmentTotals(ByVal IncludePayments As Boolean) As Double
        Dim sReturn As Double
        Dim dtTotals As DataTable

        Try
            If oFinancialCommitment.ThirdPartyPayment Then
                dtTotals = oFinancialEvent.CommitmentTotalsDatatable(2, False, False)
            Else
                dtTotals = oFinancialEvent.CommitmentTotalsDatatable(1, False, False)
            End If

            sReturn = CDbl(txtAmount.Text) - sStartingTotal
            If cmbType.Text <> "Change Order" Then
                sReturn = sReturn * -1
            End If

            Dim prevAdjustment As Double = 0.0
            If cmbType.Text = "Change Order" And strPrevType.ToUpper = "Unencumberance".ToUpper Then
                prevAdjustment = CDbl(dtTotals.Rows(0)("EventAdjustmentTotal")) * -1
                sReturn = sReturn + CDbl(dtTotals.Rows(0)("EventCommitmentTotal")) + prevAdjustment
            Else
                sReturn = sReturn + CDbl(dtTotals.Rows(0)("EventCommitmentTotal")) + CDbl(dtTotals.Rows(0)("EventAdjustmentTotal"))
            End If

            If IncludePayments Then
                sReturn -= CDbl(dtTotals.Rows(0)("EventPaymentTotal"))
            End If

            Return sReturn
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Save Calculation Failed " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Function

End Class
