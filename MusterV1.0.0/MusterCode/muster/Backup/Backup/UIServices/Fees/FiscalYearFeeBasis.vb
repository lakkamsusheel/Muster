Public Class FiscalYearFeeBasis
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Friend FeeBasisID As Int64
    Friend TemplateID As Int64
    Friend Mode As String

    Private oFeeBasis As New MUSTER.BusinessLogic.pFeeBasis

    Dim bolLoading As Boolean
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
    Friend WithEvents pnlFYBasisBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlFYBasisDetails As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents lblInvoiceAdviseGen As System.Windows.Forms.Label
    Friend WithEvents dtPickDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents lblTime As System.Windows.Forms.Label
    Friend WithEvents lblPer As System.Windows.Forms.Label
    Friend WithEvents lblAs As System.Windows.Forms.Label
    Friend WithEvents lblDollarsPer As System.Windows.Forms.Label
    Friend WithEvents lblLateFee As System.Windows.Forms.Label
    Friend WithEvents txtLateFee As System.Windows.Forms.TextBox
    Friend WithEvents lblBaseFee As System.Windows.Forms.Label
    Friend WithEvents lblLateGrace As System.Windows.Forms.Label
    Friend WithEvents lblEarlyGrace As System.Windows.Forms.Label
    Friend WithEvents dtPickLateGrace As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickEarlyGrace As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblThrough As System.Windows.Forms.Label
    Friend WithEvents txtBaseFee As System.Windows.Forms.TextBox
    Friend WithEvents lblForPeriod As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents lblCaption As System.Windows.Forms.Label
    Friend WithEvents pblFYBasisDescription As System.Windows.Forms.Panel
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents pnlFyBasis As System.Windows.Forms.Panel
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents lblEndDate As System.Windows.Forms.Label
    Friend WithEvents cmbFeeType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbLateFeePeriod As System.Windows.Forms.ComboBox
    Friend WithEvents cmbFeeUnits As System.Windows.Forms.ComboBox
    Friend WithEvents dtPickTime As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlFYBasisBottom = New System.Windows.Forms.Panel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlFYBasisDetails = New System.Windows.Forms.Panel
        Me.dtPickTime = New System.Windows.Forms.DateTimePicker
        Me.pblFYBasisDescription = New System.Windows.Forms.Panel
        Me.lblDescription = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.pnlFyBasis = New System.Windows.Forms.Panel
        Me.lblEndDate = New System.Windows.Forms.Label
        Me.lblStartDate = New System.Windows.Forms.Label
        Me.lblPer = New System.Windows.Forms.Label
        Me.lblAs = New System.Windows.Forms.Label
        Me.cmbFeeType = New System.Windows.Forms.ComboBox
        Me.cmbLateFeePeriod = New System.Windows.Forms.ComboBox
        Me.cmbFeeUnits = New System.Windows.Forms.ComboBox
        Me.lblDollarsPer = New System.Windows.Forms.Label
        Me.lblLateFee = New System.Windows.Forms.Label
        Me.txtLateFee = New System.Windows.Forms.TextBox
        Me.lblBaseFee = New System.Windows.Forms.Label
        Me.lblLateGrace = New System.Windows.Forms.Label
        Me.lblEarlyGrace = New System.Windows.Forms.Label
        Me.dtPickLateGrace = New System.Windows.Forms.DateTimePicker
        Me.dtPickEarlyGrace = New System.Windows.Forms.DateTimePicker
        Me.lblThrough = New System.Windows.Forms.Label
        Me.txtBaseFee = New System.Windows.Forms.TextBox
        Me.lblForPeriod = New System.Windows.Forms.Label
        Me.lblYear = New System.Windows.Forms.Label
        Me.lblCaption = New System.Windows.Forms.Label
        Me.lblTime = New System.Windows.Forms.Label
        Me.lblDate = New System.Windows.Forms.Label
        Me.dtPickDate = New System.Windows.Forms.DateTimePicker
        Me.lblInvoiceAdviseGen = New System.Windows.Forms.Label
        Me.pnlFYBasisBottom.SuspendLayout()
        Me.pnlFYBasisDetails.SuspendLayout()
        Me.pblFYBasisDescription.SuspendLayout()
        Me.pnlFyBasis.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlFYBasisBottom
        '
        Me.pnlFYBasisBottom.Controls.Add(Me.btnDelete)
        Me.pnlFYBasisBottom.Controls.Add(Me.btnCancel)
        Me.pnlFYBasisBottom.Controls.Add(Me.btnSave)
        Me.pnlFYBasisBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFYBasisBottom.Location = New System.Drawing.Point(0, 302)
        Me.pnlFYBasisBottom.Name = "pnlFYBasisBottom"
        Me.pnlFYBasisBottom.Size = New System.Drawing.Size(504, 40)
        Me.pnlFYBasisBottom.TabIndex = 15
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(280, 8)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.TabIndex = 18
        Me.btnDelete.Text = "Delete"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(200, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 17
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(120, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 16
        Me.btnSave.Text = "Save"
        '
        'pnlFYBasisDetails
        '
        Me.pnlFYBasisDetails.AutoScroll = True
        Me.pnlFYBasisDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlFYBasisDetails.Controls.Add(Me.dtPickTime)
        Me.pnlFYBasisDetails.Controls.Add(Me.pblFYBasisDescription)
        Me.pnlFYBasisDetails.Controls.Add(Me.pnlFyBasis)
        Me.pnlFYBasisDetails.Controls.Add(Me.lblTime)
        Me.pnlFYBasisDetails.Controls.Add(Me.lblDate)
        Me.pnlFYBasisDetails.Controls.Add(Me.dtPickDate)
        Me.pnlFYBasisDetails.Controls.Add(Me.lblInvoiceAdviseGen)
        Me.pnlFYBasisDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFYBasisDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlFYBasisDetails.Name = "pnlFYBasisDetails"
        Me.pnlFYBasisDetails.Size = New System.Drawing.Size(504, 302)
        Me.pnlFYBasisDetails.TabIndex = 10
        '
        'dtPickTime
        '
        Me.dtPickTime.Checked = False
        Me.dtPickTime.Format = System.Windows.Forms.DateTimePickerFormat.Time
        Me.dtPickTime.Location = New System.Drawing.Point(272, 187)
        Me.dtPickTime.Name = "dtPickTime"
        Me.dtPickTime.ShowUpDown = True
        Me.dtPickTime.Size = New System.Drawing.Size(96, 20)
        Me.dtPickTime.TabIndex = 14
        '
        'pblFYBasisDescription
        '
        Me.pblFYBasisDescription.Controls.Add(Me.lblDescription)
        Me.pblFYBasisDescription.Controls.Add(Me.txtDescription)
        Me.pblFYBasisDescription.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pblFYBasisDescription.Location = New System.Drawing.Point(0, 210)
        Me.pblFYBasisDescription.Name = "pblFYBasisDescription"
        Me.pblFYBasisDescription.Size = New System.Drawing.Size(500, 88)
        Me.pblFYBasisDescription.TabIndex = 13
        '
        'lblDescription
        '
        Me.lblDescription.Location = New System.Drawing.Point(32, 8)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(64, 17)
        Me.lblDescription.TabIndex = 28
        Me.lblDescription.Text = "Description"
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(96, 8)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDescription.Size = New System.Drawing.Size(368, 72)
        Me.txtDescription.TabIndex = 15
        Me.txtDescription.Text = ""
        '
        'pnlFyBasis
        '
        Me.pnlFyBasis.Controls.Add(Me.lblEndDate)
        Me.pnlFyBasis.Controls.Add(Me.lblStartDate)
        Me.pnlFyBasis.Controls.Add(Me.lblPer)
        Me.pnlFyBasis.Controls.Add(Me.lblAs)
        Me.pnlFyBasis.Controls.Add(Me.cmbFeeType)
        Me.pnlFyBasis.Controls.Add(Me.cmbLateFeePeriod)
        Me.pnlFyBasis.Controls.Add(Me.cmbFeeUnits)
        Me.pnlFyBasis.Controls.Add(Me.lblDollarsPer)
        Me.pnlFyBasis.Controls.Add(Me.lblLateFee)
        Me.pnlFyBasis.Controls.Add(Me.txtLateFee)
        Me.pnlFyBasis.Controls.Add(Me.lblBaseFee)
        Me.pnlFyBasis.Controls.Add(Me.lblLateGrace)
        Me.pnlFyBasis.Controls.Add(Me.lblEarlyGrace)
        Me.pnlFyBasis.Controls.Add(Me.dtPickLateGrace)
        Me.pnlFyBasis.Controls.Add(Me.dtPickEarlyGrace)
        Me.pnlFyBasis.Controls.Add(Me.lblThrough)
        Me.pnlFyBasis.Controls.Add(Me.txtBaseFee)
        Me.pnlFyBasis.Controls.Add(Me.lblForPeriod)
        Me.pnlFyBasis.Controls.Add(Me.lblYear)
        Me.pnlFyBasis.Controls.Add(Me.lblCaption)
        Me.pnlFyBasis.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFyBasis.Location = New System.Drawing.Point(0, 0)
        Me.pnlFyBasis.Name = "pnlFyBasis"
        Me.pnlFyBasis.Size = New System.Drawing.Size(500, 156)
        Me.pnlFyBasis.TabIndex = 0
        '
        'lblEndDate
        '
        Me.lblEndDate.BackColor = System.Drawing.SystemColors.Window
        Me.lblEndDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEndDate.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblEndDate.Location = New System.Drawing.Point(304, 40)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(64, 23)
        Me.lblEndDate.TabIndex = 42
        Me.lblEndDate.Text = "06/30/2001"
        Me.lblEndDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblStartDate
        '
        Me.lblStartDate.BackColor = System.Drawing.SystemColors.Window
        Me.lblStartDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStartDate.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblStartDate.Location = New System.Drawing.Point(128, 40)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(64, 23)
        Me.lblStartDate.TabIndex = 41
        Me.lblStartDate.Text = "07/01/2000"
        Me.lblStartDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPer
        '
        Me.lblPer.Location = New System.Drawing.Point(354, 130)
        Me.lblPer.Name = "lblPer"
        Me.lblPer.Size = New System.Drawing.Size(22, 17)
        Me.lblPer.TabIndex = 40
        Me.lblPer.Text = "per"
        '
        'lblAs
        '
        Me.lblAs.Location = New System.Drawing.Point(175, 130)
        Me.lblAs.Name = "lblAs"
        Me.lblAs.Size = New System.Drawing.Size(19, 17)
        Me.lblAs.TabIndex = 39
        Me.lblAs.Text = "as"
        '
        'cmbFeeType
        '
        Me.cmbFeeType.Location = New System.Drawing.Point(199, 130)
        Me.cmbFeeType.Name = "cmbFeeType"
        Me.cmbFeeType.Size = New System.Drawing.Size(152, 21)
        Me.cmbFeeType.TabIndex = 8
        '
        'cmbLateFeePeriod
        '
        Me.cmbLateFeePeriod.Location = New System.Drawing.Point(382, 130)
        Me.cmbLateFeePeriod.Name = "cmbLateFeePeriod"
        Me.cmbLateFeePeriod.Size = New System.Drawing.Size(112, 21)
        Me.cmbLateFeePeriod.TabIndex = 9
        '
        'cmbFeeUnits
        '
        Me.cmbFeeUnits.Location = New System.Drawing.Point(382, 106)
        Me.cmbFeeUnits.Name = "cmbFeeUnits"
        Me.cmbFeeUnits.Size = New System.Drawing.Size(112, 21)
        Me.cmbFeeUnits.TabIndex = 6
        '
        'lblDollarsPer
        '
        Me.lblDollarsPer.Location = New System.Drawing.Point(231, 106)
        Me.lblDollarsPer.Name = "lblDollarsPer"
        Me.lblDollarsPer.Size = New System.Drawing.Size(64, 17)
        Me.lblDollarsPer.TabIndex = 35
        Me.lblDollarsPer.Text = "dollars per"
        '
        'lblLateFee
        '
        Me.lblLateFee.Location = New System.Drawing.Point(7, 130)
        Me.lblLateFee.Name = "lblLateFee"
        Me.lblLateFee.Size = New System.Drawing.Size(64, 17)
        Me.lblLateFee.TabIndex = 34
        Me.lblLateFee.Text = "Late Fee"
        '
        'txtLateFee
        '
        Me.txtLateFee.Location = New System.Drawing.Point(71, 130)
        Me.txtLateFee.Name = "txtLateFee"
        Me.txtLateFee.TabIndex = 7
        Me.txtLateFee.Text = ""
        '
        'lblBaseFee
        '
        Me.lblBaseFee.Location = New System.Drawing.Point(7, 106)
        Me.lblBaseFee.Name = "lblBaseFee"
        Me.lblBaseFee.Size = New System.Drawing.Size(64, 17)
        Me.lblBaseFee.TabIndex = 32
        Me.lblBaseFee.Text = "Base Fee"
        '
        'lblLateGrace
        '
        Me.lblLateGrace.Location = New System.Drawing.Point(224, 66)
        Me.lblLateGrace.Name = "lblLateGrace"
        Me.lblLateGrace.Size = New System.Drawing.Size(64, 17)
        Me.lblLateGrace.TabIndex = 31
        Me.lblLateGrace.Text = "Late Grace"
        '
        'lblEarlyGrace
        '
        Me.lblEarlyGrace.Location = New System.Drawing.Point(63, 66)
        Me.lblEarlyGrace.Name = "lblEarlyGrace"
        Me.lblEarlyGrace.Size = New System.Drawing.Size(64, 17)
        Me.lblEarlyGrace.TabIndex = 30
        Me.lblEarlyGrace.Text = "Early Grace"
        '
        'dtPickLateGrace
        '
        Me.dtPickLateGrace.Checked = False
        Me.dtPickLateGrace.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickLateGrace.Location = New System.Drawing.Point(303, 66)
        Me.dtPickLateGrace.Name = "dtPickLateGrace"
        Me.dtPickLateGrace.Size = New System.Drawing.Size(89, 20)
        Me.dtPickLateGrace.TabIndex = 4
        '
        'dtPickEarlyGrace
        '
        Me.dtPickEarlyGrace.Checked = False
        Me.dtPickEarlyGrace.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickEarlyGrace.Location = New System.Drawing.Point(127, 66)
        Me.dtPickEarlyGrace.Name = "dtPickEarlyGrace"
        Me.dtPickEarlyGrace.Size = New System.Drawing.Size(89, 20)
        Me.dtPickEarlyGrace.TabIndex = 3
        '
        'lblThrough
        '
        Me.lblThrough.Location = New System.Drawing.Point(232, 42)
        Me.lblThrough.Name = "lblThrough"
        Me.lblThrough.Size = New System.Drawing.Size(46, 17)
        Me.lblThrough.TabIndex = 27
        Me.lblThrough.Text = "through"
        '
        'txtBaseFee
        '
        Me.txtBaseFee.Location = New System.Drawing.Point(71, 106)
        Me.txtBaseFee.Name = "txtBaseFee"
        Me.txtBaseFee.TabIndex = 5
        Me.txtBaseFee.Text = ""
        '
        'lblForPeriod
        '
        Me.lblForPeriod.Location = New System.Drawing.Point(63, 42)
        Me.lblForPeriod.Name = "lblForPeriod"
        Me.lblForPeriod.Size = New System.Drawing.Size(64, 17)
        Me.lblForPeriod.TabIndex = 23
        Me.lblForPeriod.Text = "For Period"
        '
        'lblYear
        '
        Me.lblYear.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear.Location = New System.Drawing.Point(280, 10)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.TabIndex = 22
        '
        'lblCaption
        '
        Me.lblCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCaption.Location = New System.Drawing.Point(120, 10)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(160, 23)
        Me.lblCaption.TabIndex = 21
        Me.lblCaption.Text = "Fee Basis for Fiscal Year"
        '
        'lblTime
        '
        Me.lblTime.Location = New System.Drawing.Point(240, 187)
        Me.lblTime.Name = "lblTime"
        Me.lblTime.Size = New System.Drawing.Size(32, 17)
        Me.lblTime.TabIndex = 24
        Me.lblTime.Text = "Time"
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(56, 187)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(32, 17)
        Me.lblDate.TabIndex = 23
        Me.lblDate.Text = "Date"
        '
        'dtPickDate
        '
        Me.dtPickDate.Checked = False
        Me.dtPickDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickDate.Location = New System.Drawing.Point(96, 187)
        Me.dtPickDate.Name = "dtPickDate"
        Me.dtPickDate.Size = New System.Drawing.Size(88, 20)
        Me.dtPickDate.TabIndex = 11
        '
        'lblInvoiceAdviseGen
        '
        Me.lblInvoiceAdviseGen.Location = New System.Drawing.Point(174, 163)
        Me.lblInvoiceAdviseGen.Name = "lblInvoiceAdviseGen"
        Me.lblInvoiceAdviseGen.Size = New System.Drawing.Size(152, 17)
        Me.lblInvoiceAdviseGen.TabIndex = 21
        Me.lblInvoiceAdviseGen.Text = "Invoice Advice Generation"
        '
        'FiscalYearFeeBasis
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(504, 342)
        Me.Controls.Add(Me.pnlFYBasisDetails)
        Me.Controls.Add(Me.pnlFYBasisBottom)
        Me.Name = "FiscalYearFeeBasis"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add/Modify/Delete Fiscal Year Fee Basis"
        Me.pnlFYBasisBottom.ResumeLayout(False)
        Me.pnlFYBasisDetails.ResumeLayout(False)
        Me.pblFYBasisDescription.ResumeLayout(False)
        Me.pnlFyBasis.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Support Routines"
    Private Sub LoadFeeBasisForm(ByVal nFeeBasisID As Int64)
        Dim xFeeBasis As New MUSTER.BusinessLogic.pFeeBasis

        xFeeBasis.Retrieve(nFeeBasisID)
        txtBaseFee.Text = FormatNumber(xFeeBasis.BaseFee, 2, TriState.True, TriState.False, TriState.True)
        cmbFeeUnits.SelectedValue = xFeeBasis.BaseUnit

        txtLateFee.Text = FormatNumber(xFeeBasis.LateFee, 2, TriState.True, TriState.False, TriState.True)
        cmbFeeType.SelectedValue = xFeeBasis.LateType
        cmbLateFeePeriod.SelectedValue = xFeeBasis.LatePeriod

        If Mode = "ADD" Then

            lblYear.Text = xFeeBasis.FiscalYear + 1
            lblStartDate.Text = DateAdd(DateInterval.Year, 1, xFeeBasis.PeriodStart)
            lblEndDate.Text = DateAdd(DateInterval.Year, 1, xFeeBasis.PeriodEnd)
            dtPickEarlyGrace.Value = DateAdd(DateInterval.Year, 1, xFeeBasis.EarlyGrace)
            dtPickLateGrace.Value = DateAdd(DateInterval.Year, 1, xFeeBasis.LateGrace)

            dtPickTime.Value = DateAdd(DateInterval.Hour, 2, Now)
            txtDescription.Text = ""

            oFeeBasis.BaseFee = CSng(txtBaseFee.Text)
            oFeeBasis.BaseUnit = cmbFeeUnits.SelectedValue
            oFeeBasis.Description = ""
            oFeeBasis.EarlyGrace = dtPickEarlyGrace.Value.Date
            oFeeBasis.FiscalYear = lblYear.Text
            oFeeBasis.GenerateDate = dtPickDate.Value.Date
            oFeeBasis.GenerateTime = dtPickTime.Value
            oFeeBasis.LateFee = CSng(txtLateFee.Text)
            oFeeBasis.LateGrace = dtPickLateGrace.Value.Date
            oFeeBasis.LatePeriod = cmbLateFeePeriod.SelectedValue
            oFeeBasis.LateType = cmbFeeType.SelectedValue
            oFeeBasis.PeriodEnd = lblEndDate.Text
            oFeeBasis.PeriodStart = lblStartDate.Text

        Else
            If xFeeBasis.Generated Then
                btnDelete.Enabled = False
            End If
            lblYear.Text = xFeeBasis.FiscalYear
            lblStartDate.Text = xFeeBasis.PeriodStart
            lblEndDate.Text = xFeeBasis.PeriodEnd
            dtPickEarlyGrace.Value = xFeeBasis.EarlyGrace
            dtPickLateGrace.Value = xFeeBasis.LateGrace

            dtPickDate.Value = xFeeBasis.GenerateDate.Date
            If Date.Compare(xFeeBasis.GenerateTime, CDate("01/01/0001")) = 0 Then
                dtPickTime.Value = DateAdd(DateInterval.Hour, 2, Now) 'DatePart(DateInterval.Hour, xFeeBasis.GenerateDate) & ":" & Format(DatePart(DateInterval.Minute, xFeeBasis.GenerateDate), "00")
                xFeeBasis.GenerateTime = DateAdd(DateInterval.Hour, 2, Now)
            Else
                dtPickTime.Value = xFeeBasis.GenerateTime 'DatePart(DateInterval.Hour, xFeeBasis.GenerateDate) & ":" & Format(DatePart(DateInterval.Minute, xFeeBasis.GenerateDate), "00")
            End If
            txtDescription.Text = xFeeBasis.Description
        End If

    End Sub
    Private Sub LoadDropDowns()
        Try
            Dim dtFeeUnits As DataTable = oFeeBasis.PopulateFeeUnits
            Dim dtLateFeePeriod As DataTable = oFeeBasis.PopulateLateFeePeriod
            Dim dtFeeType As DataTable = oFeeBasis.PopulateLateFeeType

            cmbFeeUnits.DataSource = dtFeeUnits
            cmbFeeUnits.DisplayMember = "PROPERTY_NAME"
            cmbFeeUnits.ValueMember = "PROPERTY_ID"

            cmbLateFeePeriod.DataSource = dtLateFeePeriod
            cmbLateFeePeriod.DisplayMember = "PROPERTY_NAME"
            cmbLateFeePeriod.ValueMember = "PROPERTY_ID"

            cmbFeeType.DataSource = dtFeeType
            cmbFeeType.DisplayMember = "PROPERTY_NAME"
            cmbFeeType.ValueMember = "PROPERTY_ID"


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try

    End Sub
#End Region
#Region "UI Control Events"
    Private Sub FiscalYearFeeBasis_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bolLoading = True

        LoadDropDowns()

        Select Case Mode
            Case "ADD"
                If TemplateID = 0 Then
                    TemplateID = oFeeBasis.GetMaxFeeBasisID
                End If
                oFeeBasis.Retrieve(0)
                LoadFeeBasisForm(TemplateID)

            Case "MODIFY"
                oFeeBasis.Retrieve(FeeBasisID)
                LoadFeeBasisForm(FeeBasisID)

            Case "REGENERATE"
                FeeBasisID = oFeeBasis.GetMaxFeeBasisID
                oFeeBasis.Retrieve(FeeBasisID)
                LoadFeeBasisForm(FeeBasisID)
                dtPickEarlyGrace.Enabled = False
                dtPickLateGrace.Enabled = False
                txtBaseFee.ReadOnly = True
                txtLateFee.ReadOnly = True
                cmbFeeType.Enabled = False
                cmbFeeUnits.Enabled = False
                cmbLateFeePeriod.Enabled = False

            Case "DELETE"
                oFeeBasis.Retrieve(FeeBasisID)
                LoadFeeBasisForm(FeeBasisID)
                btnSave.Enabled = False

        End Select

        bolLoading = False

    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub dtPickEarlyGrace_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickEarlyGrace.ValueChanged
        If bolLoading Then Exit Sub
        oFeeBasis.EarlyGrace = dtPickEarlyGrace.Value.Date

    End Sub

    Private Sub dtPickLateGrace_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickLateGrace.ValueChanged
        If bolLoading Then Exit Sub

        oFeeBasis.LateGrace = dtPickLateGrace.Value.Date
    End Sub

    Private Sub txtBaseFee_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBaseFee.TextChanged
        If bolLoading Then Exit Sub
        If IsNumeric(txtBaseFee.Text) Then
            oFeeBasis.BaseFee = txtBaseFee.Text
        Else
            oFeeBasis.BaseFee = 0
            txtBaseFee.Text = "0.00"
        End If

    End Sub

    Private Sub txtLateFee_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLateFee.TextChanged
        If bolLoading Then Exit Sub
        If IsNumeric(txtLateFee.Text) Then
            oFeeBasis.LateFee = txtLateFee.Text
        Else
            oFeeBasis.LateFee = 0
            txtLateFee.Text = "0"
        End If
    End Sub

    Private Sub cmbFeeType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFeeType.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oFeeBasis.LateType = cmbFeeType.SelectedValue

    End Sub

    Private Sub cmbFeeUnits_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFeeUnits.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oFeeBasis.BaseUnit = cmbFeeUnits.SelectedValue

    End Sub

    Private Sub cmbLateFeePeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLateFeePeriod.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oFeeBasis.LatePeriod = cmbLateFeePeriod.SelectedValue

    End Sub

    Private Sub dtPickDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickDate.ValueChanged
        If bolLoading Then Exit Sub

        oFeeBasis.GenerateDate = dtPickDate.Value
    End Sub

    Private Sub txtDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescription.TextChanged
        If bolLoading Then Exit Sub

        oFeeBasis.Description = txtDescription.Text
    End Sub

    Private Sub txtBaseFee_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBaseFee.LostFocus
        If bolLoading Then Exit Sub

        bolLoading = True

        txtBaseFee.Text = FormatNumber(txtBaseFee.Text, 2, TriState.True)

        bolLoading = False
    End Sub

    Private Sub txtLateFee_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLateFee.LostFocus
        If bolLoading Then Exit Sub

        bolLoading = True
        txtLateFee.Text = FormatNumber(txtLateFee.Text, 2, TriState.True)
        bolLoading = False
    End Sub
    Private Sub dtPickTime_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickTime.ValueChanged
        If bolLoading Then Exit Sub

        oFeeBasis.GenerateTime = dtPickTime.Value
    End Sub
#End Region


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Try

            If oFeeBasis.EarlyGrace < Now.Date Then
                MsgBox("Early Grace Date Must be Greater Than " & Now.Date.ToShortDateString)
                Exit Sub
            End If

            If oFeeBasis.LateGrace < Now.Date Then
                MsgBox("Late Grace Date Must be Greater Than " & Now.Date.ToShortDateString)
                Exit Sub
            End If

            If oFeeBasis.GenerateDate < Now.Date Then
                MsgBox("Invoice Advice Generation Date Must be Greater Than or Equal to " & Now.Date.ToShortDateString)
                Exit Sub
            End If

            If oFeeBasis.ID <= 0 Then
                oFeeBasis.CreatedBy = MusterContainer.AppUser.ID
            Else
                oFeeBasis.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oFeeBasis.Save(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            Select Case Mode
                Case "ADD"
                    ProcessCalendarEvents()
                Case "REGENERATE"
                    oFeeBasis.PurgePendingInvoiceHeaders()
                    oFeeBasis.MarkCalendarCompleted_ByDesc("Annual Billing was generated")

            End Select

            MsgBox("Fee Basis Added/Updated")

            Me.Close()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Save Changes " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub ProcessCalendarEvents()
        Dim ocalendar As New MUSTER.BusinessLogic.pCalendar

        Dim dtNotificationDate As Date = Now()
        Dim dtDueDate As Date = "05/01/" & oFeeBasis.FiscalYear
        Dim nColorCode
        Dim strTaskDesc = "Create New Fiscal Year Fees Basis"
        Dim strUserID As String = ""
        Dim strSourceUserID As String = "SYSTEM"
        Dim strGroupID As String = "Fee Admin"
        Dim bolDuetoMe As Boolean = False
        Dim bolToDo As Boolean = True
        Dim bolCompleted As Boolean = False
        Dim bolDeleted As Boolean = False

        Dim oCalendarInfo As MUSTER.Info.CalendarInfo

        Try

            oFeeBasis.MarkCalendarCompleted_ByDesc("Create New Fiscal Year Fees Basis")

            oCalendarInfo = New MUSTER.Info.CalendarInfo(0, _
                                            dtNotificationDate, _
                                            dtDueDate, _
                                            nColorCode, _
                                            strTaskDesc, _
                                            strUserID, _
                                            strSourceUserID, _
                                            strGroupID, _
                                            bolDuetoMe, _
                                            bolToDo, _
                                            bolCompleted, _
                                            bolDeleted, _
                                            "sdfsdf", _
                                            Now(), _
                                            "asdf", _
                                            Now())

            oCalendarInfo.OwningEntityID = oFeeBasis.ID
            oCalendarInfo.OwningEntityType = 25
            oCalendarInfo.IsDirty = True
            ocalendar.Add(oCalendarInfo)
            ocalendar.Flush()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Process Calendar " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try

            If MsgBox("Do you wish to delete this Fee Basis?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                Exit Sub
            End If
            oFeeBasis.Deleted = True
            If oFeeBasis.ID <= 0 Then
                oFeeBasis.CreatedBy = MusterContainer.AppUser.ID
            Else
                oFeeBasis.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oFeeBasis.Save(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            MsgBox("Fee Basis Deleted")
            Me.Close()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Delete Fee Basis " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
