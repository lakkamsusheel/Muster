Public Class AdvanceSearchSummary
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
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
    Friend WithEvents pnlContainer As System.Windows.Forms.Panel
    Friend WithEvents pnlOwnerDetails As System.Windows.Forms.Panel
    Friend WithEvents ugOwner As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblOwnerHeader As System.Windows.Forms.Label
    Friend WithEvents pnlContactdetails As System.Windows.Forms.Panel
    Friend WithEvents lblContactHeader As System.Windows.Forms.Label
    Friend WithEvents lblContactsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlFacilitiesDetails As System.Windows.Forms.Panel
    Friend WithEvents lblFacilitiesHeader As System.Windows.Forms.Label
    Friend WithEvents lblFacilitiesDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerHeader As System.Windows.Forms.Panel
    Friend WithEvents lblOwnerDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlCompanydetails As System.Windows.Forms.Panel
    Friend WithEvents lblCompanyHeader As System.Windows.Forms.Label
    Friend WithEvents lblCompanyDisplay As System.Windows.Forms.Label
    Friend WithEvents ugContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlCompanyHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlFacilitiesHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlContactHeader As System.Windows.Forms.Panel
    Friend WithEvents lblContractorHeader As System.Windows.Forms.Label
    Friend WithEvents pnlContractorDetails As System.Windows.Forms.Panel
    Friend WithEvents lblContractorDisplay As System.Windows.Forms.Label
    Friend WithEvents ugContractors As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugCompanies As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugFacilities As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlContractorHeader As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.pnlContainer = New System.Windows.Forms.Panel
        Me.pnlContractorDetails = New System.Windows.Forms.Panel
        Me.ugContractors = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlContractorHeader = New System.Windows.Forms.Panel
        Me.lblContractorHeader = New System.Windows.Forms.Label
        Me.lblContractorDisplay = New System.Windows.Forms.Label
        Me.pnlCompanydetails = New System.Windows.Forms.Panel
        Me.ugCompanies = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlCompanyHeader = New System.Windows.Forms.Panel
        Me.lblCompanyHeader = New System.Windows.Forms.Label
        Me.lblCompanyDisplay = New System.Windows.Forms.Label
        Me.pnlContactdetails = New System.Windows.Forms.Panel
        Me.ugContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlContactHeader = New System.Windows.Forms.Panel
        Me.lblContactHeader = New System.Windows.Forms.Label
        Me.lblContactsDisplay = New System.Windows.Forms.Label
        Me.pnlFacilitiesDetails = New System.Windows.Forms.Panel
        Me.ugFacilities = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFacilitiesHeader = New System.Windows.Forms.Panel
        Me.lblFacilitiesHeader = New System.Windows.Forms.Label
        Me.lblFacilitiesDisplay = New System.Windows.Forms.Label
        Me.pnlOwnerDetails = New System.Windows.Forms.Panel
        Me.ugOwner = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerHeader = New System.Windows.Forms.Panel
        Me.lblOwnerHeader = New System.Windows.Forms.Label
        Me.lblOwnerDisplay = New System.Windows.Forms.Label
        Me.pnlContainer.SuspendLayout()
        Me.pnlContractorDetails.SuspendLayout()
        CType(Me.ugContractors, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlContractorHeader.SuspendLayout()
        Me.pnlCompanydetails.SuspendLayout()
        CType(Me.ugCompanies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCompanyHeader.SuspendLayout()
        Me.pnlContactdetails.SuspendLayout()
        CType(Me.ugContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlContactHeader.SuspendLayout()
        Me.pnlFacilitiesDetails.SuspendLayout()
        CType(Me.ugFacilities, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilitiesHeader.SuspendLayout()
        Me.pnlOwnerDetails.SuspendLayout()
        CType(Me.ugOwner, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlContainer
        '
        Me.pnlContainer.AutoScroll = True
        Me.pnlContainer.Controls.Add(Me.pnlContractorDetails)
        Me.pnlContainer.Controls.Add(Me.pnlContractorHeader)
        Me.pnlContainer.Controls.Add(Me.pnlCompanydetails)
        Me.pnlContainer.Controls.Add(Me.pnlCompanyHeader)
        Me.pnlContainer.Controls.Add(Me.pnlContactdetails)
        Me.pnlContainer.Controls.Add(Me.pnlContactHeader)
        Me.pnlContainer.Controls.Add(Me.pnlFacilitiesDetails)
        Me.pnlContainer.Controls.Add(Me.pnlFacilitiesHeader)
        Me.pnlContainer.Controls.Add(Me.pnlOwnerDetails)
        Me.pnlContainer.Controls.Add(Me.pnlOwnerHeader)
        Me.pnlContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlContainer.Location = New System.Drawing.Point(0, 0)
        Me.pnlContainer.Name = "pnlContainer"
        Me.pnlContainer.Size = New System.Drawing.Size(1032, 880)
        Me.pnlContainer.TabIndex = 0
        '
        'pnlContractorDetails
        '
        Me.pnlContractorDetails.Controls.Add(Me.ugContractors)
        Me.pnlContractorDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContractorDetails.Location = New System.Drawing.Point(0, 724)
        Me.pnlContractorDetails.Name = "pnlContractorDetails"
        Me.pnlContractorDetails.Size = New System.Drawing.Size(1032, 147)
        Me.pnlContractorDetails.TabIndex = 9
        '
        'ugContractors
        '
        Me.ugContractors.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugContractors.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugContractors.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugContractors.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugContractors.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugContractors.Location = New System.Drawing.Point(16, 10)
        Me.ugContractors.Name = "ugContractors"
        Me.ugContractors.Size = New System.Drawing.Size(952, 134)
        Me.ugContractors.TabIndex = 2
        '
        'pnlContractorHeader
        '
        Me.pnlContractorHeader.Controls.Add(Me.lblContractorHeader)
        Me.pnlContractorHeader.Controls.Add(Me.lblContractorDisplay)
        Me.pnlContractorHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContractorHeader.Location = New System.Drawing.Point(0, 700)
        Me.pnlContractorHeader.Name = "pnlContractorHeader"
        Me.pnlContractorHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlContractorHeader.TabIndex = 8
        '
        'lblContractorHeader
        '
        Me.lblContractorHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblContractorHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblContractorHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractorHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblContractorHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblContractorHeader.Name = "lblContractorHeader"
        Me.lblContractorHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblContractorHeader.TabIndex = 1
        Me.lblContractorHeader.Text = "Contractor"
        Me.lblContractorHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblContractorDisplay
        '
        Me.lblContractorDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblContractorDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblContractorDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContractorDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblContractorDisplay.Name = "lblContractorDisplay"
        Me.lblContractorDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblContractorDisplay.TabIndex = 0
        Me.lblContractorDisplay.Text = "-"
        Me.lblContractorDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlCompanydetails
        '
        Me.pnlCompanydetails.Controls.Add(Me.ugCompanies)
        Me.pnlCompanydetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCompanydetails.Location = New System.Drawing.Point(0, 560)
        Me.pnlCompanydetails.Name = "pnlCompanydetails"
        Me.pnlCompanydetails.Size = New System.Drawing.Size(1032, 140)
        Me.pnlCompanydetails.TabIndex = 7
        '
        'ugCompanies
        '
        Me.ugCompanies.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCompanies.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugCompanies.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugCompanies.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugCompanies.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCompanies.Location = New System.Drawing.Point(16, 8)
        Me.ugCompanies.Name = "ugCompanies"
        Me.ugCompanies.Size = New System.Drawing.Size(952, 128)
        Me.ugCompanies.TabIndex = 2
        '
        'pnlCompanyHeader
        '
        Me.pnlCompanyHeader.Controls.Add(Me.lblCompanyHeader)
        Me.pnlCompanyHeader.Controls.Add(Me.lblCompanyDisplay)
        Me.pnlCompanyHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCompanyHeader.Location = New System.Drawing.Point(0, 536)
        Me.pnlCompanyHeader.Name = "pnlCompanyHeader"
        Me.pnlCompanyHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlCompanyHeader.TabIndex = 6
        '
        'lblCompanyHeader
        '
        Me.lblCompanyHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblCompanyHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCompanyHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompanyHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCompanyHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblCompanyHeader.Name = "lblCompanyHeader"
        Me.lblCompanyHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblCompanyHeader.TabIndex = 1
        Me.lblCompanyHeader.Text = "Company"
        Me.lblCompanyHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCompanyDisplay
        '
        Me.lblCompanyDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCompanyDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCompanyDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompanyDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblCompanyDisplay.Name = "lblCompanyDisplay"
        Me.lblCompanyDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblCompanyDisplay.TabIndex = 0
        Me.lblCompanyDisplay.Text = "-"
        Me.lblCompanyDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlContactdetails
        '
        Me.pnlContactdetails.Controls.Add(Me.ugContacts)
        Me.pnlContactdetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContactdetails.Location = New System.Drawing.Point(0, 348)
        Me.pnlContactdetails.Name = "pnlContactdetails"
        Me.pnlContactdetails.Size = New System.Drawing.Size(1032, 188)
        Me.pnlContactdetails.TabIndex = 5
        '
        'ugContacts
        '
        Me.ugContacts.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.TextHAlign = Infragistics.Win.HAlign.Left
        Me.ugContacts.DisplayLayout.Appearance = Appearance1
        Me.ugContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugContacts.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugContacts.Location = New System.Drawing.Point(16, 8)
        Me.ugContacts.Name = "ugContacts"
        Me.ugContacts.Size = New System.Drawing.Size(952, 176)
        Me.ugContacts.TabIndex = 3
        '
        'pnlContactHeader
        '
        Me.pnlContactHeader.Controls.Add(Me.lblContactHeader)
        Me.pnlContactHeader.Controls.Add(Me.lblContactsDisplay)
        Me.pnlContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContactHeader.Location = New System.Drawing.Point(0, 324)
        Me.pnlContactHeader.Name = "pnlContactHeader"
        Me.pnlContactHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlContactHeader.TabIndex = 4
        '
        'lblContactHeader
        '
        Me.lblContactHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblContactHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblContactHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContactHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblContactHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblContactHeader.Name = "lblContactHeader"
        Me.lblContactHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblContactHeader.TabIndex = 1
        Me.lblContactHeader.Text = "Contact"
        Me.lblContactHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblContactsDisplay
        '
        Me.lblContactsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblContactsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblContactsDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContactsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblContactsDisplay.Name = "lblContactsDisplay"
        Me.lblContactsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblContactsDisplay.TabIndex = 0
        Me.lblContactsDisplay.Text = "-"
        Me.lblContactsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlFacilitiesDetails
        '
        Me.pnlFacilitiesDetails.Controls.Add(Me.ugFacilities)
        Me.pnlFacilitiesDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFacilitiesDetails.Location = New System.Drawing.Point(0, 186)
        Me.pnlFacilitiesDetails.Name = "pnlFacilitiesDetails"
        Me.pnlFacilitiesDetails.Size = New System.Drawing.Size(1032, 138)
        Me.pnlFacilitiesDetails.TabIndex = 3
        '
        'ugFacilities
        '
        Me.ugFacilities.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilities.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFacilities.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugFacilities.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFacilities.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFacilities.Location = New System.Drawing.Point(16, 8)
        Me.ugFacilities.Name = "ugFacilities"
        Me.ugFacilities.Size = New System.Drawing.Size(952, 128)
        Me.ugFacilities.TabIndex = 1
        '
        'pnlFacilitiesHeader
        '
        Me.pnlFacilitiesHeader.Controls.Add(Me.lblFacilitiesHeader)
        Me.pnlFacilitiesHeader.Controls.Add(Me.lblFacilitiesDisplay)
        Me.pnlFacilitiesHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFacilitiesHeader.Location = New System.Drawing.Point(0, 162)
        Me.pnlFacilitiesHeader.Name = "pnlFacilitiesHeader"
        Me.pnlFacilitiesHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlFacilitiesHeader.TabIndex = 2
        '
        'lblFacilitiesHeader
        '
        Me.lblFacilitiesHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblFacilitiesHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblFacilitiesHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitiesHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblFacilitiesHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblFacilitiesHeader.Name = "lblFacilitiesHeader"
        Me.lblFacilitiesHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblFacilitiesHeader.TabIndex = 1
        Me.lblFacilitiesHeader.Text = "Facilities"
        Me.lblFacilitiesHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFacilitiesDisplay
        '
        Me.lblFacilitiesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFacilitiesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblFacilitiesDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitiesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblFacilitiesDisplay.Name = "lblFacilitiesDisplay"
        Me.lblFacilitiesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblFacilitiesDisplay.TabIndex = 0
        Me.lblFacilitiesDisplay.Text = "-"
        Me.lblFacilitiesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlOwnerDetails
        '
        Me.pnlOwnerDetails.Controls.Add(Me.ugOwner)
        Me.pnlOwnerDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlOwnerDetails.Name = "pnlOwnerDetails"
        Me.pnlOwnerDetails.Size = New System.Drawing.Size(1032, 138)
        Me.pnlOwnerDetails.TabIndex = 1
        '
        'ugOwner
        '
        Me.ugOwner.Cursor = System.Windows.Forms.Cursors.Default
        Appearance2.TextHAlign = Infragistics.Win.HAlign.Left
        Me.ugOwner.DisplayLayout.Appearance = Appearance2
        Me.ugOwner.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugOwner.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugOwner.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugOwner.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugOwner.Location = New System.Drawing.Point(16, 6)
        Me.ugOwner.Name = "ugOwner"
        Me.ugOwner.Size = New System.Drawing.Size(952, 128)
        Me.ugOwner.TabIndex = 0
        '
        'pnlOwnerHeader
        '
        Me.pnlOwnerHeader.Controls.Add(Me.lblOwnerHeader)
        Me.pnlOwnerHeader.Controls.Add(Me.lblOwnerDisplay)
        Me.pnlOwnerHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerHeader.Name = "pnlOwnerHeader"
        Me.pnlOwnerHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlOwnerHeader.TabIndex = 0
        '
        'lblOwnerHeader
        '
        Me.lblOwnerHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblOwnerHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblOwnerHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblOwnerHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblOwnerHeader.Name = "lblOwnerHeader"
        Me.lblOwnerHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblOwnerHeader.TabIndex = 1
        Me.lblOwnerHeader.Text = "Owner"
        Me.lblOwnerHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOwnerDisplay
        '
        Me.lblOwnerDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOwnerDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblOwnerDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblOwnerDisplay.Name = "lblOwnerDisplay"
        Me.lblOwnerDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblOwnerDisplay.TabIndex = 0
        Me.lblOwnerDisplay.Text = "-"
        Me.lblOwnerDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AdvanceSearchSummary
        '
        Me.Controls.Add(Me.pnlContainer)
        Me.Name = "AdvanceSearchSummary"
        Me.Size = New System.Drawing.Size(1032, 880)
        Me.pnlContainer.ResumeLayout(False)
        Me.pnlContractorDetails.ResumeLayout(False)
        CType(Me.ugContractors, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlContractorHeader.ResumeLayout(False)
        Me.pnlCompanydetails.ResumeLayout(False)
        CType(Me.ugCompanies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCompanyHeader.ResumeLayout(False)
        Me.pnlContactdetails.ResumeLayout(False)
        CType(Me.ugContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlContactHeader.ResumeLayout(False)
        Me.pnlFacilitiesDetails.ResumeLayout(False)
        CType(Me.ugFacilities, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilitiesHeader.ResumeLayout(False)
        Me.pnlOwnerDetails.ResumeLayout(False)
        CType(Me.ugOwner, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerHeader.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub ExpandCollapse(ByRef pnl As Panel, ByRef lbl As Label)
        pnl.Visible = Not pnl.Visible
        lbl.Text = IIf(pnl.Visible, "-", "+")
    End Sub
    Private Sub lblOwnerDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblOwnerDisplay.Click
        ExpandCollapse(Me.pnlOwnerDetails, lblOwnerDisplay)
    End Sub
    Private Sub lblFacilitiesDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFacilitiesDisplay.Click
        ExpandCollapse(Me.pnlFacilitiesDetails, lblFacilitiesDisplay)
    End Sub
    Private Sub lblContactsDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblContactsDisplay.Click
        ExpandCollapse(Me.pnlContactdetails, lblContactsDisplay)
    End Sub
    Private Sub lblCompanyDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCompanyDisplay.Click
        ExpandCollapse(Me.pnlCompanydetails, lblCompanyDisplay)
    End Sub
    Private Sub lblContractorDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblContractorDisplay.Click
        ExpandCollapse(Me.pnlContractorDetails, lblContractorDisplay)
    End Sub
    Public WriteOnly Property Owner() As DataTable
        Set(ByVal Value As DataTable)
            ugOwner.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property Facilities() As DataTable
        Set(ByVal Value As DataTable)
            ugFacilities.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property Contact() As DataTable
        Set(ByVal Value As DataTable)
            ugContacts.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property Company() As DataTable
        Set(ByVal Value As DataTable)
            ugCompanies.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property Contractor() As DataTable
        Set(ByVal Value As DataTable)
            ugContractors.DataSource = Value
        End Set
    End Property

    Private Sub ugOwner_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugOwner.InitializeLayout
        Try
            If ugOwner.DataSource.rows.count > 0 Then
                'e.Layout.Bands(0).Columns("SNo").Width = 50
                e.Layout.Bands(0).Columns("SNo").Hidden = True
                e.Layout.Bands(0).Columns("Owner ID").Width = 75
                e.Layout.Bands(0).Columns("Owner Name").Width = 200
                e.Layout.Bands(0).Columns("Owner Address").Width = 250
                e.Layout.Bands(0).Columns("Owner Address").VertScrollBar = True
                e.Layout.Bands(0).Columns("Owner Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                e.Layout.Bands(0).Columns("Owner City").Width = 100
                e.Layout.Bands(0).Columns("State").Width = 40
                e.Layout.Bands(0).Columns("Owner Contact").Width = 120
                e.Layout.Bands(0).Columns("Owner Phone").Width = 90
                e.Layout.Bands(0).Columns("Owner Points").Width = 75
                e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugFacilities_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugFacilities.InitializeLayout
        Try
            If ugFacilities.DataSource.rows.count > 0 Then
                'e.Layout.Bands(0).Columns("SNo").Width = 50
                e.Layout.Bands(0).Columns("SNo").Hidden = True
                e.Layout.Bands(0).Columns("Facility ID").Width = 75
                e.Layout.Bands(0).Columns("Owner Name").Width = 150
                e.Layout.Bands(0).Columns("Facility Name").Width = 200
                e.Layout.Bands(0).Columns("Facility Address").Width = 200
                e.Layout.Bands(0).Columns("Facility Address").VertScrollBar = True
                e.Layout.Bands(0).Columns("Facility Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                e.Layout.Bands(0).Columns("Facility City").Width = 100
                e.Layout.Bands(0).Columns("Facility County").Width = 80
                e.Layout.Bands(0).Columns("Points").Width = 50
                e.Layout.Bands(0).Columns("Lust Site").Width = 35
                e.Layout.Bands(0).Columns("CIU").Width = 35
                e.Layout.Bands(0).Columns("TOS").Width = 35
                e.Layout.Bands(0).Columns("POU").Width = 35
                e.Layout.Bands(0).Columns("TOSI").Width = 35
                e.Layout.Bands(0).Columns("Facility Contact").Width = 85
                e.Layout.Bands(0).Columns("Facility Phone").Width = 90
                e.Layout.Bands(0).Columns("Owner ID").Width = 75
                e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugContacts_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugContacts.InitializeLayout
        Try
            If ugContacts.DataSource.rows.count > 0 Then
                e.Layout.Bands(0).Columns("Contact Name").Width = 150
                e.Layout.Bands(0).Columns("Contact Address").Width = 200
                e.Layout.Bands(0).Columns("Contact Address").VertScrollBar = True
                e.Layout.Bands(0).Columns("Contact Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                e.Layout.Bands(0).Columns("Contact City").Width = 100
                e.Layout.Bands(0).Columns("Contact State").Width = 80
                e.Layout.Bands(0).Columns("Contact Phone").Width = 80
                e.Layout.Bands(0).Columns("ZipCode").Width = 75
                e.Layout.Bands(0).Columns("Contact Type").Width = 150
                e.Layout.Bands(0).Columns("Contact Source").Width = 100
                e.Layout.Bands(0).Columns("Contact Source ID").Width = 85
                e.Layout.Bands(0).Columns("Contact Points").Width = 50
                e.Layout.Bands(0).Columns("Module ID").Hidden = True
                e.Layout.Bands(0).Columns("Owner ID").Hidden = True
                e.Layout.Bands(0).Columns("Facility ID").Hidden = True
                e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugCompanies_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugCompanies.InitializeLayout
        Try
            If ugCompanies.DataSource.rows.count > 0 Then
                e.Layout.Bands(0).Columns("Company Name").Width = 180
                e.Layout.Bands(0).Columns("Company Address").Width = 200
                e.Layout.Bands(0).Columns("Company Address").VertScrollBar = True
                e.Layout.Bands(0).Columns("Company Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                e.Layout.Bands(0).Columns("City").Width = 100
                e.Layout.Bands(0).Columns("State").Width = 40
                e.Layout.Bands(0).Columns("Phone").Width = 90
                e.Layout.Bands(0).Columns("Zip").Width = 45
                e.Layout.Bands(0).Columns("Installer/Closures").Width = 100
                e.Layout.Bands(0).Columns("Closures").Width = 65
                e.Layout.Bands(0).Columns("ERAC").Width = 50
                e.Layout.Bands(0).Columns("Points").Width = 40
                e.Layout.Bands(0).Columns("Company ID").Hidden = True
                e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugContractors_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugContractors.InitializeLayout
        Try
            If ugContractors.DataSource.rows.count > 0 Then
                e.Layout.Bands(0).Columns("Licensee Name").Width = 150
                e.Layout.Bands(0).Columns("Company Name").Width = 150
                e.Layout.Bands(0).Columns("Address").Width = 160
                e.Layout.Bands(0).Columns("Address").VertScrollBar = True
                e.Layout.Bands(0).Columns("Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                e.Layout.Bands(0).Columns("City").Width = 85
                e.Layout.Bands(0).Columns("State").Width = 40
                e.Layout.Bands(0).Columns("Phone").Width = 85
                e.Layout.Bands(0).Columns("Zip").Width = 45
                e.Layout.Bands(0).Columns("Status").Width = 125
                e.Layout.Bands(0).Columns("Expiration Date").Width = 85
                e.Layout.Bands(0).Columns("Points").Width = 40
                e.Layout.Bands(0).Columns("Licensee ID").Hidden = True
                e.Layout.Bands(0).Columns("Company ID").Hidden = True
                e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
