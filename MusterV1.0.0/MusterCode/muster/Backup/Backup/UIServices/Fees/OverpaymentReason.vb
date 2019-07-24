Public Class OverpaymentReason
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"
    Friend PaymentID As Int64
    Friend OwnerID As Int64

    Dim oReceipt As New MUSTER.BusinessLogic.pFeeReceipt
    Dim returnVal As String = String.Empty

#End Region
#Region "UI Support Routines"

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
    Friend WithEvents pnlOverPaymentBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlOverPaymentDetails As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblOwnerID As System.Windows.Forms.Label
    Friend WithEvents lblCheckNumber As System.Windows.Forms.Label
    Friend WithEvents lblCheckDate As System.Windows.Forms.Label
    Friend WithEvents lblIssuingCompany As System.Windows.Forms.Label
    Friend WithEvents lblTotalAmount As System.Windows.Forms.Label
    Friend WithEvents lblOverpaymentAmount As System.Windows.Forms.Label
    Friend WithEvents lblReason As System.Windows.Forms.Label
    Friend WithEvents txtReason As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerIDValue As System.Windows.Forms.Label
    Friend WithEvents lblIssuingCompanyValue As System.Windows.Forms.Label
    Friend WithEvents lblCheckNumberValue As System.Windows.Forms.Label
    Friend WithEvents lblTotalAmountValue As System.Windows.Forms.Label
    Friend WithEvents lblOverpaymentAmountValue As System.Windows.Forms.Label
    Friend WithEvents lblCheckDateValue As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlOverPaymentBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlOverPaymentDetails = New System.Windows.Forms.Panel
        Me.lblCheckDateValue = New System.Windows.Forms.Label
        Me.lblOverpaymentAmountValue = New System.Windows.Forms.Label
        Me.lblTotalAmountValue = New System.Windows.Forms.Label
        Me.lblCheckNumberValue = New System.Windows.Forms.Label
        Me.lblIssuingCompanyValue = New System.Windows.Forms.Label
        Me.lblOwnerIDValue = New System.Windows.Forms.Label
        Me.txtReason = New System.Windows.Forms.TextBox
        Me.lblReason = New System.Windows.Forms.Label
        Me.lblOverpaymentAmount = New System.Windows.Forms.Label
        Me.lblTotalAmount = New System.Windows.Forms.Label
        Me.lblIssuingCompany = New System.Windows.Forms.Label
        Me.lblCheckDate = New System.Windows.Forms.Label
        Me.lblCheckNumber = New System.Windows.Forms.Label
        Me.lblOwnerID = New System.Windows.Forms.Label
        Me.pnlOverPaymentBottom.SuspendLayout()
        Me.pnlOverPaymentDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlOverPaymentBottom
        '
        Me.pnlOverPaymentBottom.Controls.Add(Me.btnCancel)
        Me.pnlOverPaymentBottom.Controls.Add(Me.btnSave)
        Me.pnlOverPaymentBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOverPaymentBottom.Location = New System.Drawing.Point(0, 150)
        Me.pnlOverPaymentBottom.Name = "pnlOverPaymentBottom"
        Me.pnlOverPaymentBottom.Size = New System.Drawing.Size(672, 40)
        Me.pnlOverPaymentBottom.TabIndex = 3
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(336, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(240, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save"
        '
        'pnlOverPaymentDetails
        '
        Me.pnlOverPaymentDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblCheckDateValue)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblOverpaymentAmountValue)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblTotalAmountValue)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblCheckNumberValue)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblIssuingCompanyValue)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblOwnerIDValue)
        Me.pnlOverPaymentDetails.Controls.Add(Me.txtReason)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblReason)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblOverpaymentAmount)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblTotalAmount)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblIssuingCompany)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblCheckDate)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblCheckNumber)
        Me.pnlOverPaymentDetails.Controls.Add(Me.lblOwnerID)
        Me.pnlOverPaymentDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOverPaymentDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlOverPaymentDetails.Name = "pnlOverPaymentDetails"
        Me.pnlOverPaymentDetails.Size = New System.Drawing.Size(672, 150)
        Me.pnlOverPaymentDetails.TabIndex = 0
        '
        'lblCheckDateValue
        '
        Me.lblCheckDateValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCheckDateValue.Location = New System.Drawing.Point(576, 56)
        Me.lblCheckDateValue.Name = "lblCheckDateValue"
        Me.lblCheckDateValue.Size = New System.Drawing.Size(80, 23)
        Me.lblCheckDateValue.TabIndex = 14
        '
        'lblOverpaymentAmountValue
        '
        Me.lblOverpaymentAmountValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOverpaymentAmountValue.Location = New System.Drawing.Point(424, 56)
        Me.lblOverpaymentAmountValue.Name = "lblOverpaymentAmountValue"
        Me.lblOverpaymentAmountValue.Size = New System.Drawing.Size(80, 23)
        Me.lblOverpaymentAmountValue.TabIndex = 13
        '
        'lblTotalAmountValue
        '
        Me.lblTotalAmountValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTotalAmountValue.Location = New System.Drawing.Point(232, 56)
        Me.lblTotalAmountValue.Name = "lblTotalAmountValue"
        Me.lblTotalAmountValue.Size = New System.Drawing.Size(72, 23)
        Me.lblTotalAmountValue.TabIndex = 12
        '
        'lblCheckNumberValue
        '
        Me.lblCheckNumberValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCheckNumberValue.Location = New System.Drawing.Point(96, 56)
        Me.lblCheckNumberValue.Name = "lblCheckNumberValue"
        Me.lblCheckNumberValue.Size = New System.Drawing.Size(56, 23)
        Me.lblCheckNumberValue.TabIndex = 11
        '
        'lblIssuingCompanyValue
        '
        Me.lblIssuingCompanyValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblIssuingCompanyValue.Location = New System.Drawing.Point(272, 16)
        Me.lblIssuingCompanyValue.Name = "lblIssuingCompanyValue"
        Me.lblIssuingCompanyValue.Size = New System.Drawing.Size(384, 23)
        Me.lblIssuingCompanyValue.TabIndex = 10
        '
        'lblOwnerIDValue
        '
        Me.lblOwnerIDValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOwnerIDValue.Location = New System.Drawing.Point(72, 16)
        Me.lblOwnerIDValue.Name = "lblOwnerIDValue"
        Me.lblOwnerIDValue.Size = New System.Drawing.Size(56, 23)
        Me.lblOwnerIDValue.TabIndex = 9
        '
        'txtReason
        '
        Me.txtReason.Location = New System.Drawing.Point(80, 104)
        Me.txtReason.Name = "txtReason"
        Me.txtReason.Size = New System.Drawing.Size(528, 20)
        Me.txtReason.TabIndex = 8
        Me.txtReason.Text = ""
        '
        'lblReason
        '
        Me.lblReason.Location = New System.Drawing.Point(8, 101)
        Me.lblReason.Name = "lblReason"
        Me.lblReason.Size = New System.Drawing.Size(56, 23)
        Me.lblReason.TabIndex = 7
        Me.lblReason.Text = "Reason"
        '
        'lblOverpaymentAmount
        '
        Me.lblOverpaymentAmount.Location = New System.Drawing.Point(312, 56)
        Me.lblOverpaymentAmount.Name = "lblOverpaymentAmount"
        Me.lblOverpaymentAmount.Size = New System.Drawing.Size(120, 23)
        Me.lblOverpaymentAmount.TabIndex = 6
        Me.lblOverpaymentAmount.Text = "Overpayment Amount"
        '
        'lblTotalAmount
        '
        Me.lblTotalAmount.Location = New System.Drawing.Point(160, 56)
        Me.lblTotalAmount.Name = "lblTotalAmount"
        Me.lblTotalAmount.Size = New System.Drawing.Size(78, 23)
        Me.lblTotalAmount.TabIndex = 5
        Me.lblTotalAmount.Text = "Total Amount"
        '
        'lblIssuingCompany
        '
        Me.lblIssuingCompany.Location = New System.Drawing.Point(160, 16)
        Me.lblIssuingCompany.Name = "lblIssuingCompany"
        Me.lblIssuingCompany.Size = New System.Drawing.Size(104, 23)
        Me.lblIssuingCompany.TabIndex = 4
        Me.lblIssuingCompany.Text = "Issuing Company"
        '
        'lblCheckDate
        '
        Me.lblCheckDate.Location = New System.Drawing.Point(520, 56)
        Me.lblCheckDate.Name = "lblCheckDate"
        Me.lblCheckDate.Size = New System.Drawing.Size(56, 23)
        Me.lblCheckDate.TabIndex = 3
        Me.lblCheckDate.Text = "Date"
        '
        'lblCheckNumber
        '
        Me.lblCheckNumber.Location = New System.Drawing.Point(8, 56)
        Me.lblCheckNumber.Name = "lblCheckNumber"
        Me.lblCheckNumber.Size = New System.Drawing.Size(80, 23)
        Me.lblCheckNumber.TabIndex = 2
        Me.lblCheckNumber.Text = "Check Number"
        '
        'lblOwnerID
        '
        Me.lblOwnerID.Location = New System.Drawing.Point(8, 16)
        Me.lblOwnerID.Name = "lblOwnerID"
        Me.lblOwnerID.Size = New System.Drawing.Size(56, 23)
        Me.lblOwnerID.TabIndex = 1
        Me.lblOwnerID.Text = "Owner ID"
        '
        'OverpaymentReason
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(672, 190)
        Me.Controls.Add(Me.pnlOverPaymentDetails)
        Me.Controls.Add(Me.pnlOverPaymentBottom)
        Me.Name = "OverpaymentReason"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Manage Overpayment Reason"
        Me.pnlOverPaymentBottom.ResumeLayout(False)
        Me.pnlOverPaymentDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Control Events"
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
#End Region

    Private Sub OverpaymentReason_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oReceipt.Retrieve(PaymentID)

        lblOwnerIDValue.Text = OwnerID
        lblIssuingCompanyValue.Text = oReceipt.IssuingCompany
        lblCheckNumberValue.Text = oReceipt.CheckNumber
        lblCheckDateValue.Text = oReceipt.ReceiptDate.Date
        lblOverpaymentAmountValue.Text = FormatNumber(oReceipt.AmountReceived, 2, TriState.True, TriState.False, TriState.True)
        lblOverpaymentAmountValue.ForeColor = System.Drawing.Color.Red
        lblTotalAmountValue.Text = FormatNumber(oReceipt.GetCheckTotal(OwnerID, oReceipt.CheckNumber), 2, TriState.True, TriState.False, TriState.True)
        lblTotalAmountValue.ForeColor = System.Drawing.Color.Red
        If oReceipt.OverpaymentReason <> "" Then
            txtReason.Text = oReceipt.OverpaymentReason
        End If
    End Sub



    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If oReceipt.ID <= 0 Then
                oReceipt.CreatedBy = MusterContainer.AppUser.ID
            Else
                oReceipt.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oReceipt.Save(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If


            MsgBox("Overpayment Reason Saved")
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub


    Private Sub txtReason_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReason.TextChanged
        oReceipt.OverpaymentReason = txtReason.Text

    End Sub


End Class
