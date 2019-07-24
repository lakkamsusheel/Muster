Public Class GenerateDebitMemo
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Friend OwnerID As Int64
    Friend FacilityID As Int64

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
    Friend WithEvents pnlGenDebitMemoBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlGenDebitMemoDetails As System.Windows.Forms.Panel
    Friend WithEvents btnIssueDebitMemo As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblGenDebitMemo As System.Windows.Forms.Label
    Friend WithEvents lblAdviceIDValue As System.Windows.Forms.Label
    Friend WithEvents lblAdviceID As System.Windows.Forms.Label
    Friend WithEvents lblOriginalAmtValue As System.Windows.Forms.Label
    Friend WithEvents lblOriginalAmt As System.Windows.Forms.Label
    Friend WithEvents lblOrigInvoiceValue As System.Windows.Forms.Label
    Friend WithEvents lblOrigInvoice As System.Windows.Forms.Label
    Friend WithEvents pnlGenDebitMemoHeader As System.Windows.Forms.Panel
    Friend WithEvents ugGenDebitMemo As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblOriginalInvoice As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlGenDebitMemoBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnIssueDebitMemo = New System.Windows.Forms.Button
        Me.pnlGenDebitMemoDetails = New System.Windows.Forms.Panel
        Me.ugGenDebitMemo = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlGenDebitMemoHeader = New System.Windows.Forms.Panel
        Me.lblAdviceIDValue = New System.Windows.Forms.Label
        Me.lblAdviceID = New System.Windows.Forms.Label
        Me.lblOriginalAmtValue = New System.Windows.Forms.Label
        Me.lblOriginalAmt = New System.Windows.Forms.Label
        Me.lblOrigInvoiceValue = New System.Windows.Forms.Label
        Me.lblOrigInvoice = New System.Windows.Forms.Label
        Me.lblGenDebitMemo = New System.Windows.Forms.Label
        Me.lblOriginalInvoice = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.pnlGenDebitMemoBottom.SuspendLayout()
        Me.pnlGenDebitMemoDetails.SuspendLayout()
        CType(Me.ugGenDebitMemo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlGenDebitMemoHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlGenDebitMemoBottom
        '
        Me.pnlGenDebitMemoBottom.Controls.Add(Me.btnCancel)
        Me.pnlGenDebitMemoBottom.Controls.Add(Me.btnIssueDebitMemo)
        Me.pnlGenDebitMemoBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlGenDebitMemoBottom.Location = New System.Drawing.Point(0, 414)
        Me.pnlGenDebitMemoBottom.Name = "pnlGenDebitMemoBottom"
        Me.pnlGenDebitMemoBottom.Size = New System.Drawing.Size(744, 40)
        Me.pnlGenDebitMemoBottom.TabIndex = 2
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(376, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        '
        'btnIssueDebitMemo
        '
        Me.btnIssueDebitMemo.Location = New System.Drawing.Point(264, 8)
        Me.btnIssueDebitMemo.Name = "btnIssueDebitMemo"
        Me.btnIssueDebitMemo.Size = New System.Drawing.Size(104, 23)
        Me.btnIssueDebitMemo.TabIndex = 3
        Me.btnIssueDebitMemo.Text = "Issue Debit Memo"
        '
        'pnlGenDebitMemoDetails
        '
        Me.pnlGenDebitMemoDetails.Controls.Add(Me.ugGenDebitMemo)
        Me.pnlGenDebitMemoDetails.Controls.Add(Me.pnlGenDebitMemoHeader)
        Me.pnlGenDebitMemoDetails.Controls.Add(Me.lblGenDebitMemo)
        Me.pnlGenDebitMemoDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlGenDebitMemoDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlGenDebitMemoDetails.Name = "pnlGenDebitMemoDetails"
        Me.pnlGenDebitMemoDetails.Size = New System.Drawing.Size(744, 414)
        Me.pnlGenDebitMemoDetails.TabIndex = 0
        '
        'ugGenDebitMemo
        '
        Me.ugGenDebitMemo.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugGenDebitMemo.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugGenDebitMemo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugGenDebitMemo.Location = New System.Drawing.Point(0, 72)
        Me.ugGenDebitMemo.Name = "ugGenDebitMemo"
        Me.ugGenDebitMemo.Size = New System.Drawing.Size(744, 342)
        Me.ugGenDebitMemo.TabIndex = 5
        '
        'pnlGenDebitMemoHeader
        '
        Me.pnlGenDebitMemoHeader.Controls.Add(Me.lblOriginalInvoice)
        Me.pnlGenDebitMemoHeader.Controls.Add(Me.Label2)
        Me.pnlGenDebitMemoHeader.Controls.Add(Me.lblAdviceIDValue)
        Me.pnlGenDebitMemoHeader.Controls.Add(Me.lblAdviceID)
        Me.pnlGenDebitMemoHeader.Controls.Add(Me.lblOriginalAmtValue)
        Me.pnlGenDebitMemoHeader.Controls.Add(Me.lblOriginalAmt)
        Me.pnlGenDebitMemoHeader.Controls.Add(Me.lblOrigInvoiceValue)
        Me.pnlGenDebitMemoHeader.Controls.Add(Me.lblOrigInvoice)
        Me.pnlGenDebitMemoHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlGenDebitMemoHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlGenDebitMemoHeader.Name = "pnlGenDebitMemoHeader"
        Me.pnlGenDebitMemoHeader.Size = New System.Drawing.Size(744, 72)
        Me.pnlGenDebitMemoHeader.TabIndex = 4
        '
        'lblAdviceIDValue
        '
        Me.lblAdviceIDValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAdviceIDValue.Location = New System.Drawing.Point(368, 16)
        Me.lblAdviceIDValue.Name = "lblAdviceIDValue"
        Me.lblAdviceIDValue.TabIndex = 10
        Me.lblAdviceIDValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAdviceID
        '
        Me.lblAdviceID.Location = New System.Drawing.Point(272, 16)
        Me.lblAdviceID.Name = "lblAdviceID"
        Me.lblAdviceID.TabIndex = 9
        Me.lblAdviceID.Text = "Advice Number"
        '
        'lblOriginalAmtValue
        '
        Me.lblOriginalAmtValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOriginalAmtValue.Location = New System.Drawing.Point(120, 40)
        Me.lblOriginalAmtValue.Name = "lblOriginalAmtValue"
        Me.lblOriginalAmtValue.Size = New System.Drawing.Size(96, 23)
        Me.lblOriginalAmtValue.TabIndex = 4
        Me.lblOriginalAmtValue.Text = "0.00"
        Me.lblOriginalAmtValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOriginalAmt
        '
        Me.lblOriginalAmt.Location = New System.Drawing.Point(8, 40)
        Me.lblOriginalAmt.Name = "lblOriginalAmt"
        Me.lblOriginalAmt.Size = New System.Drawing.Size(112, 23)
        Me.lblOriginalAmt.TabIndex = 3
        Me.lblOriginalAmt.Text = "Credit Memo Amount"
        '
        'lblOrigInvoiceValue
        '
        Me.lblOrigInvoiceValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOrigInvoiceValue.Location = New System.Drawing.Point(120, 16)
        Me.lblOrigInvoiceValue.Name = "lblOrigInvoiceValue"
        Me.lblOrigInvoiceValue.TabIndex = 2
        Me.lblOrigInvoiceValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOrigInvoice
        '
        Me.lblOrigInvoice.Location = New System.Drawing.Point(8, 16)
        Me.lblOrigInvoice.Name = "lblOrigInvoice"
        Me.lblOrigInvoice.TabIndex = 1
        Me.lblOrigInvoice.Text = "CM Invoice"
        '
        'lblGenDebitMemo
        '
        Me.lblGenDebitMemo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGenDebitMemo.Location = New System.Drawing.Point(120, 8)
        Me.lblGenDebitMemo.Name = "lblGenDebitMemo"
        Me.lblGenDebitMemo.Size = New System.Drawing.Size(136, 23)
        Me.lblGenDebitMemo.TabIndex = 0
        Me.lblGenDebitMemo.Text = "Generate Debit Memo"
        '
        'lblOriginalInvoice
        '
        Me.lblOriginalInvoice.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOriginalInvoice.Location = New System.Drawing.Point(624, 16)
        Me.lblOriginalInvoice.Name = "lblOriginalInvoice"
        Me.lblOriginalInvoice.TabIndex = 12
        Me.lblOriginalInvoice.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(512, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Original  Invoice"
        '
        'GenerateDebitMemo
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(744, 454)
        Me.Controls.Add(Me.pnlGenDebitMemoDetails)
        Me.Controls.Add(Me.pnlGenDebitMemoBottom)
        Me.Name = "GenerateDebitMemo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Generate Debit Memo"
        Me.pnlGenDebitMemoBottom.ResumeLayout(False)
        Me.pnlGenDebitMemoDetails.ResumeLayout(False)
        CType(Me.ugGenDebitMemo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlGenDebitMemoHeader.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Control Events"
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
#End Region

    Private Sub GenerateDebitMemo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadugGenDebitMemo()

    End Sub

    Private Sub LoadugGenDebitMemo()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16

        Try

            dsLocal = oFeeInvoice.GetCreditMemos_ByOwnerID(OwnerID)
            ugGenDebitMemo.DataSource = dsLocal

            ugGenDebitMemo.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugGenDebitMemo.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            ugGenDebitMemo.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows

                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("InvoiceID").Hidden = True

                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("CreditTotal").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("CreditTotal").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("CreditTotal").Header.Caption = "Credit Total"
                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("CreditMemoInvoice").Header.Caption = "Credit Memo Invoice"
                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("CreditMemoDate").Header.Caption = "Credit Memo Date"
                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("BillingInvoice").Header.Caption = "Billing Invoice"
                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("CreditTotal").CellAppearance.ForeColor = System.Drawing.Color.Red

                ugGenDebitMemo.DisplayLayout.Bands(0).SortedColumns.Clear()
                ugGenDebitMemo.DisplayLayout.Bands(1).SortedColumns.Clear()
                ugGenDebitMemo.DisplayLayout.Bands(0).Columns("CreditMemoInvoice").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("Facility_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("InvoiceID").Hidden = True
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("CreditMemoInvoice").Hidden = True
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("LineAmount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("LineAmount").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("LineAmount").CellAppearance.ForeColor = System.Drawing.Color.Red
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("LineAmount").Header.Caption = "Line Amount"
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("Facility_ID").Header.Caption = "Facility ID"
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("FacilityName").Header.Caption = "Facility Name"
                ugGenDebitMemo.DisplayLayout.Bands(1).Columns("FacilityName").ColSpan = 2

                ugGenDebitMemo.ActiveRow = ugGenDebitMemo.Rows(0)
                UpdateHeader()
            Else
                MsgBox("No Eligible Credit Memos Available For Credit Reversal.")
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub


    Private Sub ugGenDebitMemo_AfterRowActivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugGenDebitMemo.AfterRowActivate
        Try

            UpdateHeader()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub UpdateHeader()
        Try
            If ugGenDebitMemo.ActiveRow.Band.Index = 0 Then

                oInvoice.Retrieve(ugGenDebitMemo.ActiveRow.Cells("InvoiceID").Value)
                lblOrigInvoiceValue.Text = oInvoice.WarrantNumber
                lblOriginalAmtValue.Text = FormatNumber(oInvoice.InvoiceAmount, 2, TriState.True, TriState.False, TriState.True)
                lblOriginalInvoice.Text = oInvoice.CreditApplyTo
                lblAdviceIDValue.Text = oInvoice.InvoiceAdviceID

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnIssueDebitMemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIssueDebitMemo.Click

        Try
            oInvoice.GenerateDebitMemo(oInvoice.ID, CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            MsgBox("Debit Memo Created")
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub


End Class
