Public Class OwnerSummary
    Inherits System.Windows.Forms.UserControl



    'changes Made
    '2-12-2009  Thomas Franey      Ciber   Function:  ugPreviousFacs_InitializeLayout
    '                                      Test:      tested with owner 221 9which throws the erroe) and 12536 which does not 
    '                                      Diagnosis: When it has previous facility records, it throws an error when looking for column address
    '                                                 does not exist due to some change.          
    '                                      Solution:  Applied Existance check on columns in Previous facilites due to possible change in column headers.
    '                                                 Also applied .NET standard writing of function
    '                                      Long-term Solution Idea:  Create a Gridhandler class to handle all modifications
    '                                                                to a grid's column & row details, in which will also do checks
    '                                                                on an existance of a column or row.   


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
    Friend WithEvents pnlLustSitesHeader As System.Windows.Forms.Panel
    Friend WithEvents lblLustSitesDisplay As System.Windows.Forms.Label
    Friend WithEvents lblLustSitesHeader As System.Windows.Forms.Label
    Friend WithEvents pnlLustSiteDetails As System.Windows.Forms.Panel
    Friend WithEvents ugLustSites As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlFinancialHeader As System.Windows.Forms.Panel
    Friend WithEvents lblFinancialHeader As System.Windows.Forms.Label
    Friend WithEvents lblFinancialDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlFinancialDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlFeesHeader As System.Windows.Forms.Panel
    Friend WithEvents lblFeesHeader As System.Windows.Forms.Label
    Friend WithEvents lblFeesDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlFeesdetails As System.Windows.Forms.Panel
    Friend WithEvents pnlPenalitiesHeader As System.Windows.Forms.Panel
    Friend WithEvents lblPenalitiesHeader As System.Windows.Forms.Label
    Friend WithEvents lblPenalitiesDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlPenalitiesdetails As System.Windows.Forms.Panel
    Friend WithEvents pnlPreviousFacHeader As System.Windows.Forms.Panel
    Friend WithEvents lblPreviousFacHeader As System.Windows.Forms.Label
    Friend WithEvents lblPreviousFacDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlPreviousFacDetails As System.Windows.Forms.Panel
    Friend WithEvents ugFees As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugFinancialSites As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugPreviousFacs As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugPenalities As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlCommitmentsTotals As System.Windows.Forms.Panel
    Friend WithEvents lblToDateBalance As System.Windows.Forms.Label
    Friend WithEvents lblCurrentAdjustments As System.Windows.Forms.Label
    Friend WithEvents lblCurrentCredits As System.Windows.Forms.Label
    Friend WithEvents lblCurrentPayments As System.Windows.Forms.Label
    Friend WithEvents lblTotalDue As System.Windows.Forms.Label
    Friend WithEvents lblLateFees As System.Windows.Forms.Label
    Friend WithEvents lblCurrentFees As System.Windows.Forms.Label
    Friend WithEvents lblPriorBalance As System.Windows.Forms.Label
    Friend WithEvents lblTotals As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.pnlContainer = New System.Windows.Forms.Panel
        Me.pnlPreviousFacDetails = New System.Windows.Forms.Panel
        Me.ugPreviousFacs = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlPreviousFacHeader = New System.Windows.Forms.Panel
        Me.lblPreviousFacHeader = New System.Windows.Forms.Label
        Me.lblPreviousFacDisplay = New System.Windows.Forms.Label
        Me.pnlPenalitiesdetails = New System.Windows.Forms.Panel
        Me.ugPenalities = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlPenalitiesHeader = New System.Windows.Forms.Panel
        Me.lblPenalitiesHeader = New System.Windows.Forms.Label
        Me.lblPenalitiesDisplay = New System.Windows.Forms.Label
        Me.pnlFeesdetails = New System.Windows.Forms.Panel
        Me.pnlCommitmentsTotals = New System.Windows.Forms.Panel
        Me.lblToDateBalance = New System.Windows.Forms.Label
        Me.lblCurrentAdjustments = New System.Windows.Forms.Label
        Me.lblCurrentCredits = New System.Windows.Forms.Label
        Me.lblCurrentPayments = New System.Windows.Forms.Label
        Me.lblTotalDue = New System.Windows.Forms.Label
        Me.lblLateFees = New System.Windows.Forms.Label
        Me.lblCurrentFees = New System.Windows.Forms.Label
        Me.lblPriorBalance = New System.Windows.Forms.Label
        Me.lblTotals = New System.Windows.Forms.Label
        Me.ugFees = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFeesHeader = New System.Windows.Forms.Panel
        Me.lblFeesHeader = New System.Windows.Forms.Label
        Me.lblFeesDisplay = New System.Windows.Forms.Label
        Me.pnlFinancialDetails = New System.Windows.Forms.Panel
        Me.ugFinancialSites = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFinancialHeader = New System.Windows.Forms.Panel
        Me.lblFinancialHeader = New System.Windows.Forms.Label
        Me.lblFinancialDisplay = New System.Windows.Forms.Label
        Me.pnlLustSiteDetails = New System.Windows.Forms.Panel
        Me.ugLustSites = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlLustSitesHeader = New System.Windows.Forms.Panel
        Me.lblLustSitesHeader = New System.Windows.Forms.Label
        Me.lblLustSitesDisplay = New System.Windows.Forms.Label
        Me.pnlContainer.SuspendLayout()
        Me.pnlPreviousFacDetails.SuspendLayout()
        CType(Me.ugPreviousFacs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPreviousFacHeader.SuspendLayout()
        Me.pnlPenalitiesdetails.SuspendLayout()
        CType(Me.ugPenalities, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPenalitiesHeader.SuspendLayout()
        Me.pnlFeesdetails.SuspendLayout()
        Me.pnlCommitmentsTotals.SuspendLayout()
        CType(Me.ugFees, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFeesHeader.SuspendLayout()
        Me.pnlFinancialDetails.SuspendLayout()
        CType(Me.ugFinancialSites, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFinancialHeader.SuspendLayout()
        Me.pnlLustSiteDetails.SuspendLayout()
        CType(Me.ugLustSites, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLustSitesHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlContainer
        '
        Me.pnlContainer.AutoScroll = True
        Me.pnlContainer.Controls.Add(Me.pnlPreviousFacDetails)
        Me.pnlContainer.Controls.Add(Me.pnlPreviousFacHeader)
        Me.pnlContainer.Controls.Add(Me.pnlPenalitiesdetails)
        Me.pnlContainer.Controls.Add(Me.pnlPenalitiesHeader)
        Me.pnlContainer.Controls.Add(Me.pnlFeesdetails)
        Me.pnlContainer.Controls.Add(Me.pnlFeesHeader)
        Me.pnlContainer.Controls.Add(Me.pnlFinancialDetails)
        Me.pnlContainer.Controls.Add(Me.pnlFinancialHeader)
        Me.pnlContainer.Controls.Add(Me.pnlLustSiteDetails)
        Me.pnlContainer.Controls.Add(Me.pnlLustSitesHeader)
        Me.pnlContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlContainer.Location = New System.Drawing.Point(0, 0)
        Me.pnlContainer.Name = "pnlContainer"
        Me.pnlContainer.Size = New System.Drawing.Size(1032, 880)
        Me.pnlContainer.TabIndex = 0
        '
        'pnlPreviousFacDetails
        '
        Me.pnlPreviousFacDetails.Controls.Add(Me.ugPreviousFacs)
        Me.pnlPreviousFacDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPreviousFacDetails.Location = New System.Drawing.Point(0, 724)
        Me.pnlPreviousFacDetails.Name = "pnlPreviousFacDetails"
        Me.pnlPreviousFacDetails.Size = New System.Drawing.Size(1032, 147)
        Me.pnlPreviousFacDetails.TabIndex = 9
        '
        'ugPreviousFacs
        '
        Me.ugPreviousFacs.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPreviousFacs.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugPreviousFacs.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugPreviousFacs.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugPreviousFacs.Location = New System.Drawing.Point(16, 10)
        Me.ugPreviousFacs.Name = "ugPreviousFacs"
        Me.ugPreviousFacs.Size = New System.Drawing.Size(824, 128)
        Me.ugPreviousFacs.TabIndex = 2
        '
        'pnlPreviousFacHeader
        '
        Me.pnlPreviousFacHeader.Controls.Add(Me.lblPreviousFacHeader)
        Me.pnlPreviousFacHeader.Controls.Add(Me.lblPreviousFacDisplay)
        Me.pnlPreviousFacHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPreviousFacHeader.Location = New System.Drawing.Point(0, 700)
        Me.pnlPreviousFacHeader.Name = "pnlPreviousFacHeader"
        Me.pnlPreviousFacHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlPreviousFacHeader.TabIndex = 8
        '
        'lblPreviousFacHeader
        '
        Me.lblPreviousFacHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPreviousFacHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPreviousFacHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPreviousFacHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPreviousFacHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblPreviousFacHeader.Name = "lblPreviousFacHeader"
        Me.lblPreviousFacHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblPreviousFacHeader.TabIndex = 1
        Me.lblPreviousFacHeader.Text = "Previous Facilities"
        Me.lblPreviousFacHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPreviousFacDisplay
        '
        Me.lblPreviousFacDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPreviousFacDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPreviousFacDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPreviousFacDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPreviousFacDisplay.Name = "lblPreviousFacDisplay"
        Me.lblPreviousFacDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPreviousFacDisplay.TabIndex = 0
        Me.lblPreviousFacDisplay.Text = "-"
        Me.lblPreviousFacDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPenalitiesdetails
        '
        Me.pnlPenalitiesdetails.Controls.Add(Me.ugPenalities)
        Me.pnlPenalitiesdetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPenalitiesdetails.Location = New System.Drawing.Point(0, 560)
        Me.pnlPenalitiesdetails.Name = "pnlPenalitiesdetails"
        Me.pnlPenalitiesdetails.Size = New System.Drawing.Size(1032, 140)
        Me.pnlPenalitiesdetails.TabIndex = 7
        '
        'ugPenalities
        '
        Me.ugPenalities.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPenalities.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugPenalities.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugPenalities.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugPenalities.Location = New System.Drawing.Point(16, 8)
        Me.ugPenalities.Name = "ugPenalities"
        Me.ugPenalities.Size = New System.Drawing.Size(824, 128)
        Me.ugPenalities.TabIndex = 2
        '
        'pnlPenalitiesHeader
        '
        Me.pnlPenalitiesHeader.Controls.Add(Me.lblPenalitiesHeader)
        Me.pnlPenalitiesHeader.Controls.Add(Me.lblPenalitiesDisplay)
        Me.pnlPenalitiesHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPenalitiesHeader.Location = New System.Drawing.Point(0, 536)
        Me.pnlPenalitiesHeader.Name = "pnlPenalitiesHeader"
        Me.pnlPenalitiesHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlPenalitiesHeader.TabIndex = 6
        '
        'lblPenalitiesHeader
        '
        Me.lblPenalitiesHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPenalitiesHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPenalitiesHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPenalitiesHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPenalitiesHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblPenalitiesHeader.Name = "lblPenalitiesHeader"
        Me.lblPenalitiesHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblPenalitiesHeader.TabIndex = 1
        Me.lblPenalitiesHeader.Text = "Penalties"
        Me.lblPenalitiesHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPenalitiesDisplay
        '
        Me.lblPenalitiesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPenalitiesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPenalitiesDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPenalitiesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPenalitiesDisplay.Name = "lblPenalitiesDisplay"
        Me.lblPenalitiesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPenalitiesDisplay.TabIndex = 0
        Me.lblPenalitiesDisplay.Text = "-"
        Me.lblPenalitiesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlFeesdetails
        '
        Me.pnlFeesdetails.Controls.Add(Me.pnlCommitmentsTotals)
        Me.pnlFeesdetails.Controls.Add(Me.ugFees)
        Me.pnlFeesdetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFeesdetails.Location = New System.Drawing.Point(0, 348)
        Me.pnlFeesdetails.Name = "pnlFeesdetails"
        Me.pnlFeesdetails.Size = New System.Drawing.Size(1032, 188)
        Me.pnlFeesdetails.TabIndex = 5
        '
        'pnlCommitmentsTotals
        '
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblToDateBalance)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblCurrentAdjustments)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblCurrentCredits)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblCurrentPayments)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblTotalDue)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblLateFees)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblCurrentFees)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblPriorBalance)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblTotals)
        Me.pnlCommitmentsTotals.Location = New System.Drawing.Point(24, 140)
        Me.pnlCommitmentsTotals.Name = "pnlCommitmentsTotals"
        Me.pnlCommitmentsTotals.Size = New System.Drawing.Size(816, 32)
        Me.pnlCommitmentsTotals.TabIndex = 7
        '
        'lblToDateBalance
        '
        Me.lblToDateBalance.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblToDateBalance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblToDateBalance.Location = New System.Drawing.Point(670, 6)
        Me.lblToDateBalance.Name = "lblToDateBalance"
        Me.lblToDateBalance.Size = New System.Drawing.Size(86, 23)
        Me.lblToDateBalance.TabIndex = 8
        Me.lblToDateBalance.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurrentAdjustments
        '
        Me.lblCurrentAdjustments.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCurrentAdjustments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCurrentAdjustments.Location = New System.Drawing.Point(577, 6)
        Me.lblCurrentAdjustments.Name = "lblCurrentAdjustments"
        Me.lblCurrentAdjustments.Size = New System.Drawing.Size(93, 23)
        Me.lblCurrentAdjustments.TabIndex = 7
        Me.lblCurrentAdjustments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurrentCredits
        '
        Me.lblCurrentCredits.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCurrentCredits.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCurrentCredits.Location = New System.Drawing.Point(488, 6)
        Me.lblCurrentCredits.Name = "lblCurrentCredits"
        Me.lblCurrentCredits.Size = New System.Drawing.Size(89, 23)
        Me.lblCurrentCredits.TabIndex = 6
        Me.lblCurrentCredits.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurrentPayments
        '
        Me.lblCurrentPayments.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCurrentPayments.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCurrentPayments.Location = New System.Drawing.Point(392, 6)
        Me.lblCurrentPayments.Name = "lblCurrentPayments"
        Me.lblCurrentPayments.Size = New System.Drawing.Size(97, 23)
        Me.lblCurrentPayments.TabIndex = 5
        Me.lblCurrentPayments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotalDue
        '
        Me.lblTotalDue.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTotalDue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTotalDue.Location = New System.Drawing.Point(312, 6)
        Me.lblTotalDue.Name = "lblTotalDue"
        Me.lblTotalDue.Size = New System.Drawing.Size(80, 23)
        Me.lblTotalDue.TabIndex = 4
        Me.lblTotalDue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLateFees
        '
        Me.lblLateFees.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblLateFees.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLateFees.Location = New System.Drawing.Point(232, 6)
        Me.lblLateFees.Name = "lblLateFees"
        Me.lblLateFees.Size = New System.Drawing.Size(82, 23)
        Me.lblLateFees.TabIndex = 3
        Me.lblLateFees.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurrentFees
        '
        Me.lblCurrentFees.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCurrentFees.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCurrentFees.Location = New System.Drawing.Point(157, 6)
        Me.lblCurrentFees.Name = "lblCurrentFees"
        Me.lblCurrentFees.Size = New System.Drawing.Size(75, 23)
        Me.lblCurrentFees.TabIndex = 2
        Me.lblCurrentFees.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPriorBalance
        '
        Me.lblPriorBalance.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPriorBalance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPriorBalance.Location = New System.Drawing.Point(71, 6)
        Me.lblPriorBalance.Name = "lblPriorBalance"
        Me.lblPriorBalance.Size = New System.Drawing.Size(86, 23)
        Me.lblPriorBalance.TabIndex = 1
        Me.lblPriorBalance.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotals
        '
        Me.lblTotals.Location = New System.Drawing.Point(8, 8)
        Me.lblTotals.Name = "lblTotals"
        Me.lblTotals.Size = New System.Drawing.Size(48, 17)
        Me.lblTotals.TabIndex = 0
        Me.lblTotals.Text = "Totals:"
        '
        'ugFees
        '
        Me.ugFees.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.TextHAlign = Infragistics.Win.HAlign.Left
        Me.ugFees.DisplayLayout.Appearance = Appearance1
        Me.ugFees.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFees.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugFees.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFees.Location = New System.Drawing.Point(16, 8)
        Me.ugFees.Name = "ugFees"
        Me.ugFees.Size = New System.Drawing.Size(824, 128)
        Me.ugFees.TabIndex = 3
        '
        'pnlFeesHeader
        '
        Me.pnlFeesHeader.Controls.Add(Me.lblFeesHeader)
        Me.pnlFeesHeader.Controls.Add(Me.lblFeesDisplay)
        Me.pnlFeesHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFeesHeader.Location = New System.Drawing.Point(0, 324)
        Me.pnlFeesHeader.Name = "pnlFeesHeader"
        Me.pnlFeesHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlFeesHeader.TabIndex = 4
        '
        'lblFeesHeader
        '
        Me.lblFeesHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblFeesHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblFeesHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFeesHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblFeesHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblFeesHeader.Name = "lblFeesHeader"
        Me.lblFeesHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblFeesHeader.TabIndex = 1
        Me.lblFeesHeader.Text = "Fees"
        Me.lblFeesHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFeesDisplay
        '
        Me.lblFeesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFeesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblFeesDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFeesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblFeesDisplay.Name = "lblFeesDisplay"
        Me.lblFeesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblFeesDisplay.TabIndex = 0
        Me.lblFeesDisplay.Text = "-"
        Me.lblFeesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlFinancialDetails
        '
        Me.pnlFinancialDetails.Controls.Add(Me.ugFinancialSites)
        Me.pnlFinancialDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFinancialDetails.Location = New System.Drawing.Point(0, 186)
        Me.pnlFinancialDetails.Name = "pnlFinancialDetails"
        Me.pnlFinancialDetails.Size = New System.Drawing.Size(1032, 138)
        Me.pnlFinancialDetails.TabIndex = 3
        '
        'ugFinancialSites
        '
        Me.ugFinancialSites.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFinancialSites.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFinancialSites.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugFinancialSites.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFinancialSites.Location = New System.Drawing.Point(16, 8)
        Me.ugFinancialSites.Name = "ugFinancialSites"
        Me.ugFinancialSites.Size = New System.Drawing.Size(824, 128)
        Me.ugFinancialSites.TabIndex = 1
        '
        'pnlFinancialHeader
        '
        Me.pnlFinancialHeader.Controls.Add(Me.lblFinancialHeader)
        Me.pnlFinancialHeader.Controls.Add(Me.lblFinancialDisplay)
        Me.pnlFinancialHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFinancialHeader.Location = New System.Drawing.Point(0, 162)
        Me.pnlFinancialHeader.Name = "pnlFinancialHeader"
        Me.pnlFinancialHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlFinancialHeader.TabIndex = 2
        '
        'lblFinancialHeader
        '
        Me.lblFinancialHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblFinancialHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblFinancialHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFinancialHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblFinancialHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblFinancialHeader.Name = "lblFinancialHeader"
        Me.lblFinancialHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblFinancialHeader.TabIndex = 1
        Me.lblFinancialHeader.Text = "Financial Sites"
        Me.lblFinancialHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFinancialDisplay
        '
        Me.lblFinancialDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFinancialDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblFinancialDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFinancialDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblFinancialDisplay.Name = "lblFinancialDisplay"
        Me.lblFinancialDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblFinancialDisplay.TabIndex = 0
        Me.lblFinancialDisplay.Text = "-"
        Me.lblFinancialDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlLustSiteDetails
        '
        Me.pnlLustSiteDetails.Controls.Add(Me.ugLustSites)
        Me.pnlLustSiteDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLustSiteDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlLustSiteDetails.Name = "pnlLustSiteDetails"
        Me.pnlLustSiteDetails.Size = New System.Drawing.Size(1032, 138)
        Me.pnlLustSiteDetails.TabIndex = 1
        '
        'ugLustSites
        '
        Me.ugLustSites.Cursor = System.Windows.Forms.Cursors.Default
        Appearance2.TextHAlign = Infragistics.Win.HAlign.Left
        Me.ugLustSites.DisplayLayout.Appearance = Appearance2
        Me.ugLustSites.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugLustSites.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugLustSites.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugLustSites.Location = New System.Drawing.Point(16, 6)
        Me.ugLustSites.Name = "ugLustSites"
        Me.ugLustSites.Size = New System.Drawing.Size(824, 128)
        Me.ugLustSites.TabIndex = 0
        '
        'pnlLustSitesHeader
        '
        Me.pnlLustSitesHeader.Controls.Add(Me.lblLustSitesHeader)
        Me.pnlLustSitesHeader.Controls.Add(Me.lblLustSitesDisplay)
        Me.pnlLustSitesHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLustSitesHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlLustSitesHeader.Name = "pnlLustSitesHeader"
        Me.pnlLustSitesHeader.Size = New System.Drawing.Size(1032, 24)
        Me.pnlLustSitesHeader.TabIndex = 0
        '
        'lblLustSitesHeader
        '
        Me.lblLustSitesHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblLustSitesHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblLustSitesHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLustSitesHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblLustSitesHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblLustSitesHeader.Name = "lblLustSitesHeader"
        Me.lblLustSitesHeader.Size = New System.Drawing.Size(1016, 24)
        Me.lblLustSitesHeader.TabIndex = 1
        Me.lblLustSitesHeader.Text = "Lust Sites"
        Me.lblLustSitesHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblLustSitesDisplay
        '
        Me.lblLustSitesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLustSitesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblLustSitesDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLustSitesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblLustSitesDisplay.Name = "lblLustSitesDisplay"
        Me.lblLustSitesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblLustSitesDisplay.TabIndex = 0
        Me.lblLustSitesDisplay.Text = "-"
        Me.lblLustSitesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'OwnerSummary
        '
        Me.Controls.Add(Me.pnlContainer)
        Me.Name = "OwnerSummary"
        Me.Size = New System.Drawing.Size(1032, 880)
        Me.pnlContainer.ResumeLayout(False)
        Me.pnlPreviousFacDetails.ResumeLayout(False)
        CType(Me.ugPreviousFacs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPreviousFacHeader.ResumeLayout(False)
        Me.pnlPenalitiesdetails.ResumeLayout(False)
        CType(Me.ugPenalities, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPenalitiesHeader.ResumeLayout(False)
        Me.pnlFeesdetails.ResumeLayout(False)
        Me.pnlCommitmentsTotals.ResumeLayout(False)
        CType(Me.ugFees, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFeesHeader.ResumeLayout(False)
        Me.pnlFinancialDetails.ResumeLayout(False)
        CType(Me.ugFinancialSites, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFinancialHeader.ResumeLayout(False)
        Me.pnlLustSiteDetails.ResumeLayout(False)
        CType(Me.ugLustSites, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLustSitesHeader.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ExpandCollapse(ByRef pnl As Panel, ByRef lbl As Label)
        pnl.Visible = Not pnl.Visible
        lbl.Text = IIf(pnl.Visible, "-", "+")
    End Sub
    Private Sub lblFeesDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblFeesDisplay.Click
        ExpandCollapse(Me.pnlFeesdetails, lblFeesDisplay)
    End Sub
    Private Sub lblFinancialDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblFinancialDisplay.Click
        ExpandCollapse(Me.pnlFinancialDetails, lblFinancialDisplay)
    End Sub
    Private Sub lblLustSitesDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblLustSitesDisplay.Click
        ExpandCollapse(Me.pnlLustSiteDetails, lblLustSitesDisplay)
    End Sub
    Private Sub lblPenalitiesDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPenalitiesDisplay.Click
        ExpandCollapse(Me.pnlPenalitiesdetails, lblPenalitiesDisplay)
    End Sub
    Private Sub lblPreviousFacDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPreviousFacDisplay.Click
        ExpandCollapse(Me.pnlPreviousFacDetails, lblPreviousFacDisplay)
    End Sub

    Private Sub AlignFeeSummaryLabels()
        Try
            ' align the summary textboxes
            Dim left As Integer = ugFees.DisplayLayout.Grid.Location.X + ugFees.DisplayLayout.Bands(0).Columns("Facility ID").Width

            lblPriorBalance.Left = left
            lblPriorBalance.Width = ugFees.DisplayLayout.Bands(0).Columns("Prior Balance").Width
            left += lblPriorBalance.Width

            lblCurrentFees.Left = left
            lblCurrentFees.Width = ugFees.DisplayLayout.Bands(0).Columns("Current Fees").Width
            left += lblCurrentFees.Width

            lblLateFees.Left = left
            lblLateFees.Width = ugFees.DisplayLayout.Bands(0).Columns("Late Fees").Width
            left += lblLateFees.Width

            lblTotalDue.Left = left
            lblTotalDue.Width = ugFees.DisplayLayout.Bands(0).Columns("Total Due").Width
            left += lblTotalDue.Width

            lblCurrentPayments.Left = left
            lblCurrentPayments.Width = ugFees.DisplayLayout.Bands(0).Columns("Current Payments").Width
            left += lblCurrentPayments.Width

            lblCurrentCredits.Left = left
            lblCurrentCredits.Width = ugFees.DisplayLayout.Bands(0).Columns("Current Credits").Width
            left += lblCurrentCredits.Width

            lblCurrentAdjustments.Left = left
            lblCurrentAdjustments.Width = ugFees.DisplayLayout.Bands(0).Columns("Current Adjustments").Width
            left += lblCurrentAdjustments.Width

            lblToDateBalance.Left = left
            lblToDateBalance.Width = ugFees.DisplayLayout.Bands(0).Columns("To Date Balance").Width
            left += lblToDateBalance.Width
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Public WriteOnly Property LustSites() As DataTable
        Set(ByVal Value As DataTable)
            ugLustSites.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property FinancialSites() As DataTable
        Set(ByVal Value As DataTable)
            ugFinancialSites.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property Fees() As DataTable
        Set(ByVal Value As DataTable)
            ugFees.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property Penalities() As DataTable
        Set(ByVal Value As DataTable)
            ugPenalities.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property PreviousFacilities() As DataTable
        Set(ByVal Value As DataTable)
            ugPreviousFacs.DataSource = Value
        End Set
    End Property
    Public WriteOnly Property CurrentFees() As String
        Set(ByVal Value As String)
            Me.lblCurrentFees.Text = Value
        End Set
    End Property
    Public WriteOnly Property PriorBalance() As String
        Set(ByVal Value As String)
            Me.lblPriorBalance.Text = Value
        End Set
    End Property
    Public WriteOnly Property LateFees() As String
        Set(ByVal Value As String)
            Me.lblLateFees.Text = Value
        End Set
    End Property
    Public WriteOnly Property TotalDue() As String
        Set(ByVal Value As String)
            Me.lblTotalDue.Text = Value
        End Set
    End Property
    Public WriteOnly Property CurrentPayments() As String
        Set(ByVal Value As String)
            Me.lblCurrentPayments.Text = Value
        End Set
    End Property
    Public WriteOnly Property CurrentCredits() As String
        Set(ByVal Value As String)
            Me.lblCurrentCredits.Text = Value
        End Set
    End Property
    Public WriteOnly Property CurrentAdjustments() As String
        Set(ByVal Value As String)
            Me.lblCurrentAdjustments.Text = Value
        End Set
    End Property
    Public WriteOnly Property ToDateBalance() As String
        Set(ByVal Value As String)
            Me.lblToDateBalance.Text = Value
        End Set
    End Property

    Private Sub ugPreviousFacs_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugPreviousFacs.InitializeLayout
        Try

            'Detrmines if columns exist before modifying width and characteristics
            If ugPreviousFacs.DataSource.rows.count > 0 Then
                With e.Layout.Bands(0).Columns


                    .Item("FacilityID").Width = 80
                    .Item("Facility Name").Width = 150

                    If .Exists("Address") Then
                        With .Item("Address")

                            .Width = 200
                            .VertScrollBar = True
                            .AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                        End With
                    End If

                    If .Exists("City") Then
                        .Item("City").Width = 110
                    End If


                    If .Exists("Transfer_date") Then
                        .Item("Transfer Date").Width = 100
                        .Item("Transfer Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    End If

                    If .Exists("From") Then
                        .Item("From").Width = 100
                    End If

                    If .Exists("To") Then
                        .Item("To").Width = 100
                    End If

                    e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                    e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                End With

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugLustSites_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugLustSites.InitializeLayout
        Try
            If ugLustSites.DataSource.rows.count > 0 Then
                e.Layout.Bands(0).Columns("Priority").Width = 55
                e.Layout.Bands(0).Columns("Status").Width = 55
                e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFees_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugFees.InitializeLayout
        Try
            'If ugFees.DataSource.rows.count > 0 Then
            e.Layout.Bands(0).Columns("Facility ID").Width = 70
            e.Layout.Bands(0).Columns("Prior Balance").Width = 75
            e.Layout.Bands(0).Columns("Prior Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            e.Layout.Bands(0).Columns("Current Fees").Width = 75
            e.Layout.Bands(0).Columns("Current Fees").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            e.Layout.Bands(0).Columns("Late Fees").Width = 75
            e.Layout.Bands(0).Columns("Late Fees").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            e.Layout.Bands(0).Columns("Total Due").Width = 75
            e.Layout.Bands(0).Columns("Total Due").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            e.Layout.Bands(0).Columns("Current Payments").Width = 92
            e.Layout.Bands(0).Columns("Current Payments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            e.Layout.Bands(0).Columns("Current Credits").Width = 90
            e.Layout.Bands(0).Columns("Current Credits").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            e.Layout.Bands(0).Columns("To Date Balance").Width = 90
            e.Layout.Bands(0).Columns("To Date Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            e.Layout.Bands(0).Columns("Current Adjustments").Width = 100
            e.Layout.Bands(0).Columns("Current Adjustments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            e.Layout.Bands(0).Columns("Legal").Width = 50
            e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            'End If
            ' align the summary textboxes
            AlignFeeSummaryLabels()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFees_AfterColPosChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs) Handles ugFees.AfterColPosChanged
        AlignFeeSummaryLabels()
    End Sub
    Private Sub ugFinancialSites_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugFinancialSites.InitializeLayout
        Try
            If ugFinancialSites.DataSource.rows.count > 0 Then
                e.Layout.Bands(0).Columns("Commitment").MaskInput = "$nnn,nnn,nnn.nn"
                e.Layout.Bands(0).Columns("Commitment").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                e.Layout.Bands(0).Columns("Commitment").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw
                e.Layout.Bands(0).Columns("Commitment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                e.Layout.Bands(0).Columns("Adjustment").MaskInput = "$nnn,nnn,nnn.nn"
                e.Layout.Bands(0).Columns("Adjustment").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                e.Layout.Bands(0).Columns("Adjustment").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw
                e.Layout.Bands(0).Columns("Adjustment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                e.Layout.Bands(0).Columns("Balance").MaskInput = "$nnn,nnn,nnn.nn"
                e.Layout.Bands(0).Columns("Balance").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                e.Layout.Bands(0).Columns("Balance").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw
                e.Layout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                e.Layout.Bands(0).Columns("Payment").MaskInput = "$nnn,nnn,nnn.nn"
                e.Layout.Bands(0).Columns("Payment").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                e.Layout.Bands(0).Columns("Payment").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw
                e.Layout.Bands(0).Columns("Payment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                e.Layout.Bands(0).Columns("Commitment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                e.Layout.Bands(0).Columns("Adjustment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                e.Layout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                e.Layout.Bands(0).Columns("Payment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugPenalities_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugPenalities.InitializeLayout
        Try
            e.Layout.Bands(0).Columns("Policy Amount").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(0).Columns("Policy Amount").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(0).Columns("Policy Amount").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw
            e.Layout.Bands(0).Columns("Policy Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            e.Layout.Bands(0).Columns("Actual Amount").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(0).Columns("Actual Amount").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(0).Columns("Actual Amount").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw
            e.Layout.Bands(0).Columns("Actual Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            e.Layout.Bands(0).Columns("Amount Paid").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(0).Columns("Amount Paid").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(0).Columns("Amount Paid").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw
            e.Layout.Bands(0).Columns("Amount Paid").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            e.Layout.Bands(0).Columns("Date Assessed").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("Date Assessed").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            e.Layout.Bands(0).Columns("Date Due").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("Date Due").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            e.Layout.Bands(0).Columns("Date Paid").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("Date Paid").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            e.Layout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

End Class
