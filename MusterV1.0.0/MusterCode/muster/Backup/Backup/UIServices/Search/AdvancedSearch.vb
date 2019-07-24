Imports System.Threading

Public Class AdvancedSearch
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.ShowComment.vb
    '   Provides the mechanism for performing Advanced Searches.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1        JC      1/02/04    Changed AppUser.UserName to AppUser.ID to
    '                                  accomodate new use of pUser by application.
    '  1.2        EN     02/01/05     Modified and Added new functions to incorporate the AdvanceSearch Object.                              
    '  1.3        AN      02/10/05    Integrated AppFlags new object model
    '  1.4        JVC2    08/04/05    Addressed UST Bug 258 by adding call to btnClear_Click if the contents
    '                                   of cmbFavoriteSearch.Text is empty and by adding storage/retrieval of
    '                                   cmbFavoriteSearch.SelectedIndex in same.Tag on tab click.
    '-------------------------------------------------------------------------------
    Inherits System.Windows.Forms.Form
    Friend WithEvents frmRegServices As MusterContainer
    Friend WithEvents oAdvanceSearch As MUSTER.BusinessLogic.pAdvancedSearch
    Public WithEvents UCAdvSearchSummary As New MUSTER.AdvanceSearchSummary
    Public Event SearchResultSelection(ByVal OwnerID As Integer, ByVal FacilityID As Integer, ByVal Search_Type As String)
    Public Event CompanySearchSelection(ByVal CompanyID As Integer, ByVal LicenseeID As Integer, ByVal AssocID As Integer, ByVal Search_Type As String)
    Friend MyGUID As New System.Guid
    Dim bolIndexChanged As Boolean = False
    Dim nSelIndexValue As Integer = -1
    Private WithEvents frmConfirmSave As New FavSearchDialog
    Dim bolAssignValue As Boolean = False
    Dim BolNewSearch As Boolean = False
    Dim dtFavSearchChildList As DataTable
    Private nLastLookForRow As Int32 = -1
    Private bolAdvancedSearchError As Boolean = False
    Dim bolTabClickEvent As Boolean = False
    Private bolSearchTypeSelected As Boolean = False
    Private bolFavSearchSelected As Boolean = False
    Public strModule As String = String.Empty
    Public nEventID As Integer = 0
    Public nClosureID As Integer = 0
    Public Delegate Sub OwnerSearchHandler()
    Public strTankStatus As String = String.Empty
    Public nLustStatus As Integer = 0
    Dim SummaryThread As Thread


#Region "User Defined Variables"
    Dim dCol1 As New DataColumn
    Dim dCol2 As New DataColumn
    Dim dsAdvSearch As New System.Data.DataSet
    Dim dTableAdvSearch As New System.Data.DataTable("FilterTable")
    Dim dGridCellAdvSearch As DataGridCell
    Dim dGridAdvSearchRowCount As Integer = 0
    Dim tableStyle As DataGridTableStyle
    'Dim aColumnTextColumn As DataGridEnableTextBoxColumn
    Private bolIsLoaded As Boolean = False
#End Region
#Region "Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        MyGUID = System.Guid.NewGuid

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppUser.LogEntry("Advanced Search", MyGUID.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGUID.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)


    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MusterContainer.AppSemaphores.Remove(MyGUID.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
        MusterContainer.AppUser.LogExit(MyGUID.ToString)
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
    Friend WithEvents cmbFavoriteSearches As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSearchType As System.Windows.Forms.ComboBox
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents lstViewResults As System.Windows.Forms.ListView
    Friend WithEvents lblFavouriteSearch As System.Windows.Forms.Label
    Friend WithEvents lblSearchType As System.Windows.Forms.Label
    Friend WithEvents btnSaveSearch As System.Windows.Forms.Button
    Friend WithEvents lstSearchBys As System.Windows.Forms.ListBox
    Friend WithEvents lblSearchBys As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnDown As System.Windows.Forms.Button
    Friend WithEvents btnUp As System.Windows.Forms.Button
    Friend WithEvents lblSearchByLookFor As System.Windows.Forms.Label
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents tbCtrlAdvancedSearch As System.Windows.Forms.TabControl
    Friend WithEvents tbPageAdvancedSearch As System.Windows.Forms.TabPage
    Friend WithEvents tbPageAdvancedSearchResult As System.Windows.Forms.TabPage
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblSearchResult As System.Windows.Forms.Label
    Friend WithEvents grpActiveTanks As System.Windows.Forms.GroupBox
    Friend WithEvents rdActive As System.Windows.Forms.RadioButton
    Friend WithEvents rdClose As System.Windows.Forms.RadioButton
    Friend WithEvents rdActiveUnknown As System.Windows.Forms.RadioButton
    Friend WithEvents grpLUSTSite As System.Windows.Forms.GroupBox
    Friend WithEvents rdLUSTUnknown As System.Windows.Forms.RadioButton
    Friend WithEvents rdClosed As System.Windows.Forms.RadioButton
    Friend WithEvents rdOpen As System.Windows.Forms.RadioButton
    Friend WithEvents rdLUSTOpenorClosed As System.Windows.Forms.RadioButton
    Friend WithEvents pnlSearchResultTop As System.Windows.Forms.Panel
    Friend WithEvents pnlSearchResult As System.Windows.Forms.Panel
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents ugAdvSearch As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugSearchByLookFor As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents lblRecordCount As System.Windows.Forms.Label
    Friend WithEvents lblSearchCriterion As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.cmbFavoriteSearches = New System.Windows.Forms.ComboBox
        Me.lblFavouriteSearch = New System.Windows.Forms.Label
        Me.lblSearchType = New System.Windows.Forms.Label
        Me.cmbSearchType = New System.Windows.Forms.ComboBox
        Me.btnSaveSearch = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.ugSearchByLookFor = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.grpLUSTSite = New System.Windows.Forms.GroupBox
        Me.rdLUSTOpenorClosed = New System.Windows.Forms.RadioButton
        Me.rdLUSTUnknown = New System.Windows.Forms.RadioButton
        Me.rdClosed = New System.Windows.Forms.RadioButton
        Me.rdOpen = New System.Windows.Forms.RadioButton
        Me.grpActiveTanks = New System.Windows.Forms.GroupBox
        Me.rdActiveUnknown = New System.Windows.Forms.RadioButton
        Me.rdClose = New System.Windows.Forms.RadioButton
        Me.rdActive = New System.Windows.Forms.RadioButton
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnClear = New System.Windows.Forms.Button
        Me.btnSearch = New System.Windows.Forms.Button
        Me.lblSearchByLookFor = New System.Windows.Forms.Label
        Me.btnDown = New System.Windows.Forms.Button
        Me.btnUp = New System.Windows.Forms.Button
        Me.btnRemove = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.lblSearchBys = New System.Windows.Forms.Label
        Me.lstSearchBys = New System.Windows.Forms.ListBox
        Me.lstViewResults = New System.Windows.Forms.ListView
        Me.tbCtrlAdvancedSearch = New System.Windows.Forms.TabControl
        Me.tbPageAdvancedSearch = New System.Windows.Forms.TabPage
        Me.tbPageAdvancedSearchResult = New System.Windows.Forms.TabPage
        Me.pnlSearchResult = New System.Windows.Forms.Panel
        Me.ugAdvSearch = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlSearchResultTop = New System.Windows.Forms.Panel
        Me.btnPrint = New System.Windows.Forms.Button
        Me.lblSearchResult = New System.Windows.Forms.Label
        Me.lblSearchCriterion = New System.Windows.Forms.Label
        Me.lblRecordCount = New System.Windows.Forms.Label
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader10 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader11 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader12 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader13 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader14 = New System.Windows.Forms.ColumnHeader
        Me.pnlTop.SuspendLayout()
        CType(Me.ugSearchByLookFor, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpLUSTSite.SuspendLayout()
        Me.grpActiveTanks.SuspendLayout()
        Me.tbCtrlAdvancedSearch.SuspendLayout()
        Me.tbPageAdvancedSearch.SuspendLayout()
        Me.tbPageAdvancedSearchResult.SuspendLayout()
        Me.pnlSearchResult.SuspendLayout()
        CType(Me.ugAdvSearch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlSearchResultTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbFavoriteSearches
        '
        Me.cmbFavoriteSearches.Location = New System.Drawing.Point(112, 48)
        Me.cmbFavoriteSearches.Name = "cmbFavoriteSearches"
        Me.cmbFavoriteSearches.Size = New System.Drawing.Size(152, 21)
        Me.cmbFavoriteSearches.TabIndex = 1
        '
        'lblFavouriteSearch
        '
        Me.lblFavouriteSearch.Location = New System.Drawing.Point(8, 48)
        Me.lblFavouriteSearch.Name = "lblFavouriteSearch"
        Me.lblFavouriteSearch.Size = New System.Drawing.Size(104, 16)
        Me.lblFavouriteSearch.TabIndex = 1
        Me.lblFavouriteSearch.Text = "Favorite Searches:"
        '
        'lblSearchType
        '
        Me.lblSearchType.Location = New System.Drawing.Point(32, 17)
        Me.lblSearchType.Name = "lblSearchType"
        Me.lblSearchType.Size = New System.Drawing.Size(72, 16)
        Me.lblSearchType.TabIndex = 3
        Me.lblSearchType.Text = "Search Type:"
        '
        'cmbSearchType
        '
        Me.cmbSearchType.Items.AddRange(New Object() {"All", "Owner", "Facility", "Contact", "Contractor", "Company"})
        Me.cmbSearchType.Location = New System.Drawing.Point(112, 17)
        Me.cmbSearchType.Name = "cmbSearchType"
        Me.cmbSearchType.Size = New System.Drawing.Size(152, 21)
        Me.cmbSearchType.TabIndex = 0
        '
        'btnSaveSearch
        '
        Me.btnSaveSearch.Location = New System.Drawing.Point(272, 48)
        Me.btnSaveSearch.Name = "btnSaveSearch"
        Me.btnSaveSearch.Size = New System.Drawing.Size(112, 23)
        Me.btnSaveSearch.TabIndex = 13
        Me.btnSaveSearch.Text = "Save As Favorite"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(392, 48)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(96, 23)
        Me.btnDelete.TabIndex = 14
        Me.btnDelete.Text = "Delete Favorite"
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.ugSearchByLookFor)
        Me.pnlTop.Controls.Add(Me.grpLUSTSite)
        Me.pnlTop.Controls.Add(Me.grpActiveTanks)
        Me.pnlTop.Controls.Add(Me.btnClose)
        Me.pnlTop.Controls.Add(Me.btnClear)
        Me.pnlTop.Controls.Add(Me.btnSearch)
        Me.pnlTop.Controls.Add(Me.lblSearchByLookFor)
        Me.pnlTop.Controls.Add(Me.btnDown)
        Me.pnlTop.Controls.Add(Me.btnUp)
        Me.pnlTop.Controls.Add(Me.btnRemove)
        Me.pnlTop.Controls.Add(Me.btnAdd)
        Me.pnlTop.Controls.Add(Me.lblSearchBys)
        Me.pnlTop.Controls.Add(Me.lstSearchBys)
        Me.pnlTop.Controls.Add(Me.lblFavouriteSearch)
        Me.pnlTop.Controls.Add(Me.lblSearchType)
        Me.pnlTop.Controls.Add(Me.cmbSearchType)
        Me.pnlTop.Controls.Add(Me.cmbFavoriteSearches)
        Me.pnlTop.Controls.Add(Me.btnSaveSearch)
        Me.pnlTop.Controls.Add(Me.btnDelete)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(816, 500)
        Me.pnlTop.TabIndex = 31
        '
        'ugSearchByLookFor
        '
        Me.ugSearchByLookFor.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ugSearchByLookFor.DisplayLayout.Appearance = Appearance1
        Appearance2.BackColor = System.Drawing.Color.Aqua
        Appearance2.BackColor2 = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Appearance2.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel
        Appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.BackwardDiagonal
        Me.ugSearchByLookFor.DisplayLayout.Override.ActiveRowAppearance = Appearance2
        Appearance3.BackColor = System.Drawing.Color.RoyalBlue
        Me.ugSearchByLookFor.DisplayLayout.Override.HeaderAppearance = Appearance3
        Me.ugSearchByLookFor.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugSearchByLookFor.Location = New System.Drawing.Point(272, 104)
        Me.ugSearchByLookFor.Name = "ugSearchByLookFor"
        Me.ugSearchByLookFor.Size = New System.Drawing.Size(288, 192)
        Me.ugSearchByLookFor.TabIndex = 5
        '
        'grpLUSTSite
        '
        Me.grpLUSTSite.Controls.Add(Me.rdLUSTOpenorClosed)
        Me.grpLUSTSite.Controls.Add(Me.rdLUSTUnknown)
        Me.grpLUSTSite.Controls.Add(Me.rdClosed)
        Me.grpLUSTSite.Controls.Add(Me.rdOpen)
        Me.grpLUSTSite.Location = New System.Drawing.Point(16, 328)
        Me.grpLUSTSite.Name = "grpLUSTSite"
        Me.grpLUSTSite.Size = New System.Drawing.Size(120, 120)
        Me.grpLUSTSite.TabIndex = 8
        Me.grpLUSTSite.TabStop = False
        Me.grpLUSTSite.Text = "LUST Status"
        '
        'rdLUSTOpenorClosed
        '
        Me.rdLUSTOpenorClosed.Location = New System.Drawing.Point(24, 64)
        Me.rdLUSTOpenorClosed.Name = "rdLUSTOpenorClosed"
        Me.rdLUSTOpenorClosed.Size = New System.Drawing.Size(88, 24)
        Me.rdLUSTOpenorClosed.TabIndex = 2
        Me.rdLUSTOpenorClosed.Text = "Open/Closed"
        '
        'rdLUSTUnknown
        '
        Me.rdLUSTUnknown.Checked = True
        Me.rdLUSTUnknown.Location = New System.Drawing.Point(24, 88)
        Me.rdLUSTUnknown.Name = "rdLUSTUnknown"
        Me.rdLUSTUnknown.Size = New System.Drawing.Size(80, 24)
        Me.rdLUSTUnknown.TabIndex = 3
        Me.rdLUSTUnknown.TabStop = True
        Me.rdLUSTUnknown.Text = "Unknown"
        '
        'rdClosed
        '
        Me.rdClosed.Location = New System.Drawing.Point(24, 40)
        Me.rdClosed.Name = "rdClosed"
        Me.rdClosed.Size = New System.Drawing.Size(80, 24)
        Me.rdClosed.TabIndex = 1
        Me.rdClosed.Text = "Closed"
        '
        'rdOpen
        '
        Me.rdOpen.Location = New System.Drawing.Point(24, 16)
        Me.rdOpen.Name = "rdOpen"
        Me.rdOpen.Size = New System.Drawing.Size(80, 24)
        Me.rdOpen.TabIndex = 0
        Me.rdOpen.Text = "Open"
        '
        'grpActiveTanks
        '
        Me.grpActiveTanks.Controls.Add(Me.rdActiveUnknown)
        Me.grpActiveTanks.Controls.Add(Me.rdClose)
        Me.grpActiveTanks.Controls.Add(Me.rdActive)
        Me.grpActiveTanks.Location = New System.Drawing.Point(144, 328)
        Me.grpActiveTanks.Name = "grpActiveTanks"
        Me.grpActiveTanks.Size = New System.Drawing.Size(112, 96)
        Me.grpActiveTanks.TabIndex = 9
        Me.grpActiveTanks.TabStop = False
        Me.grpActiveTanks.Text = "Tank Status"
        '
        'rdActiveUnknown
        '
        Me.rdActiveUnknown.Checked = True
        Me.rdActiveUnknown.Location = New System.Drawing.Point(24, 64)
        Me.rdActiveUnknown.Name = "rdActiveUnknown"
        Me.rdActiveUnknown.Size = New System.Drawing.Size(80, 24)
        Me.rdActiveUnknown.TabIndex = 2
        Me.rdActiveUnknown.TabStop = True
        Me.rdActiveUnknown.Text = "Unknown"
        '
        'rdClose
        '
        Me.rdClose.Location = New System.Drawing.Point(24, 40)
        Me.rdClose.Name = "rdClose"
        Me.rdClose.Size = New System.Drawing.Size(80, 24)
        Me.rdClose.TabIndex = 1
        Me.rdClose.Text = "Closed"
        '
        'rdActive
        '
        Me.rdActive.Location = New System.Drawing.Point(24, 16)
        Me.rdActive.Name = "rdActive"
        Me.rdActive.Size = New System.Drawing.Size(80, 24)
        Me.rdActive.TabIndex = 0
        Me.rdActive.Text = "Active"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(512, 416)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 23)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Close"
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(416, 416)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(80, 23)
        Me.btnClear.TabIndex = 11
        Me.btnClear.Text = "Clear"
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(312, 416)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(80, 23)
        Me.btnSearch.TabIndex = 10
        Me.btnSearch.Text = "Search"
        '
        'lblSearchByLookFor
        '
        Me.lblSearchByLookFor.Location = New System.Drawing.Point(288, 80)
        Me.lblSearchByLookFor.Name = "lblSearchByLookFor"
        Me.lblSearchByLookFor.Size = New System.Drawing.Size(144, 16)
        Me.lblSearchByLookFor.TabIndex = 38
        Me.lblSearchByLookFor.Text = "Search Bys  /  Look Fors"
        '
        'btnDown
        '
        Me.btnDown.Location = New System.Drawing.Point(568, 200)
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(80, 23)
        Me.btnDown.TabIndex = 7
        Me.btnDown.Text = "Down"
        '
        'btnUp
        '
        Me.btnUp.Location = New System.Drawing.Point(568, 168)
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(80, 23)
        Me.btnUp.TabIndex = 6
        Me.btnUp.Text = "Up"
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(184, 200)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(80, 23)
        Me.btnRemove.TabIndex = 4
        Me.btnRemove.Text = "< Remove"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(184, 168)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(80, 23)
        Me.btnAdd.TabIndex = 3
        Me.btnAdd.Text = "Add >"
        '
        'lblSearchBys
        '
        Me.lblSearchBys.Location = New System.Drawing.Point(16, 88)
        Me.lblSearchBys.Name = "lblSearchBys"
        Me.lblSearchBys.Size = New System.Drawing.Size(120, 16)
        Me.lblSearchBys.TabIndex = 33
        Me.lblSearchBys.Text = "Available Search Bys"
        '
        'lstSearchBys
        '
        Me.lstSearchBys.Items.AddRange(New Object() {"All", "Brand", "Facility Address", "Facility Lat Degree", "Facility Lat Minutes", "Facility Long Degree", "Facility Long Minutes", "Facility City", "Facility County", "Facility Name", "Facility AIID", "Owner Address", "Owner City", "Owner ID", "Owner Name", "Project Manager", "Project Manager (History)", "Contact Company Name", "Contact First Name", "Contact Middle Name", "Contact Last Name", "Company Name", "Licensee First Name", "Licensee Middle Name", "Licensee Last Name", "Type of Service", "All Phone Numbers"})
        Me.lstSearchBys.Location = New System.Drawing.Point(16, 112)
        Me.lstSearchBys.Name = "lstSearchBys"
        Me.lstSearchBys.Size = New System.Drawing.Size(152, 173)
        Me.lstSearchBys.TabIndex = 2
        '
        'lstViewResults
        '
        Me.lstViewResults.Activation = System.Windows.Forms.ItemActivation.TwoClick
        Me.lstViewResults.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstViewResults.FullRowSelect = True
        Me.lstViewResults.GridLines = True
        Me.lstViewResults.HoverSelection = True
        Me.lstViewResults.Location = New System.Drawing.Point(-8, 424)
        Me.lstViewResults.Name = "lstViewResults"
        Me.lstViewResults.Size = New System.Drawing.Size(816, 112)
        Me.lstViewResults.TabIndex = 0
        Me.lstViewResults.View = System.Windows.Forms.View.Details
        '
        'tbCtrlAdvancedSearch
        '
        Me.tbCtrlAdvancedSearch.Controls.Add(Me.tbPageAdvancedSearch)
        Me.tbCtrlAdvancedSearch.Controls.Add(Me.tbPageAdvancedSearchResult)
        Me.tbCtrlAdvancedSearch.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlAdvancedSearch.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlAdvancedSearch.Name = "tbCtrlAdvancedSearch"
        Me.tbCtrlAdvancedSearch.SelectedIndex = 0
        Me.tbCtrlAdvancedSearch.Size = New System.Drawing.Size(824, 526)
        Me.tbCtrlAdvancedSearch.TabIndex = 33
        '
        'tbPageAdvancedSearch
        '
        Me.tbPageAdvancedSearch.AutoScroll = True
        Me.tbPageAdvancedSearch.Controls.Add(Me.pnlTop)
        Me.tbPageAdvancedSearch.Location = New System.Drawing.Point(4, 22)
        Me.tbPageAdvancedSearch.Name = "tbPageAdvancedSearch"
        Me.tbPageAdvancedSearch.Size = New System.Drawing.Size(816, 500)
        Me.tbPageAdvancedSearch.TabIndex = 0
        Me.tbPageAdvancedSearch.Text = "Advanced Search"
        '
        'tbPageAdvancedSearchResult
        '
        Me.tbPageAdvancedSearchResult.AutoScroll = True
        Me.tbPageAdvancedSearchResult.Controls.Add(Me.pnlSearchResult)
        Me.tbPageAdvancedSearchResult.Controls.Add(Me.pnlSearchResultTop)
        Me.tbPageAdvancedSearchResult.Location = New System.Drawing.Point(4, 22)
        Me.tbPageAdvancedSearchResult.Name = "tbPageAdvancedSearchResult"
        Me.tbPageAdvancedSearchResult.Size = New System.Drawing.Size(816, 500)
        Me.tbPageAdvancedSearchResult.TabIndex = 1
        Me.tbPageAdvancedSearchResult.Text = "Advanced Search Result "
        '
        'pnlSearchResult
        '
        Me.pnlSearchResult.Controls.Add(Me.ugAdvSearch)
        Me.pnlSearchResult.Controls.Add(Me.lstViewResults)
        Me.pnlSearchResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlSearchResult.Location = New System.Drawing.Point(0, 40)
        Me.pnlSearchResult.Name = "pnlSearchResult"
        Me.pnlSearchResult.Size = New System.Drawing.Size(816, 460)
        Me.pnlSearchResult.TabIndex = 4
        '
        'ugAdvSearch
        '
        Me.ugAdvSearch.Cursor = System.Windows.Forms.Cursors.Default
        Appearance4.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugAdvSearch.DisplayLayout.Override.CellAppearance = Appearance4
        Me.ugAdvSearch.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance5.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugAdvSearch.DisplayLayout.Override.RowAppearance = Appearance5
        Me.ugAdvSearch.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAdvSearch.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAdvSearch.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugAdvSearch.Location = New System.Drawing.Point(0, 0)
        Me.ugAdvSearch.Name = "ugAdvSearch"
        Me.ugAdvSearch.Size = New System.Drawing.Size(816, 460)
        Me.ugAdvSearch.TabIndex = 4
        '
        'pnlSearchResultTop
        '
        Me.pnlSearchResultTop.Controls.Add(Me.btnPrint)
        Me.pnlSearchResultTop.Controls.Add(Me.lblSearchResult)
        Me.pnlSearchResultTop.Controls.Add(Me.lblSearchCriterion)
        Me.pnlSearchResultTop.Controls.Add(Me.lblRecordCount)
        Me.pnlSearchResultTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlSearchResultTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlSearchResultTop.Name = "pnlSearchResultTop"
        Me.pnlSearchResultTop.Size = New System.Drawing.Size(816, 40)
        Me.pnlSearchResultTop.TabIndex = 4
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(728, 8)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "Print"
        '
        'lblSearchResult
        '
        Me.lblSearchResult.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchResult.Location = New System.Drawing.Point(8, 5)
        Me.lblSearchResult.Name = "lblSearchResult"
        Me.lblSearchResult.Size = New System.Drawing.Size(152, 16)
        Me.lblSearchResult.TabIndex = 2
        Me.lblSearchResult.Text = "Owner Search Result"
        Me.lblSearchResult.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblSearchCriterion
        '
        Me.lblSearchCriterion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchCriterion.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblSearchCriterion.Location = New System.Drawing.Point(176, 8)
        Me.lblSearchCriterion.Name = "lblSearchCriterion"
        Me.lblSearchCriterion.Size = New System.Drawing.Size(520, 24)
        Me.lblSearchCriterion.TabIndex = 3
        Me.lblSearchCriterion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblRecordCount
        '
        Me.lblRecordCount.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRecordCount.Location = New System.Drawing.Point(8, 22)
        Me.lblRecordCount.Name = "lblRecordCount"
        Me.lblRecordCount.Size = New System.Drawing.Size(152, 16)
        Me.lblRecordCount.TabIndex = 2
        Me.lblRecordCount.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "Owner Name"
        Me.ColumnHeader8.Width = 122
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Street Address"
        Me.ColumnHeader9.Width = 127
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "City"
        Me.ColumnHeader10.Width = 66
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "State"
        Me.ColumnHeader11.Width = 47
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "Zip"
        Me.ColumnHeader12.Width = 48
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "Facility ID"
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "Facility Name"
        Me.ColumnHeader14.Width = 142
        '
        'AdvancedSearch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(824, 526)
        Me.Controls.Add(Me.tbCtrlAdvancedSearch)
        Me.Name = "AdvancedSearch"
        Me.Text = "Advanced Search"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTop.ResumeLayout(False)
        CType(Me.ugSearchByLookFor, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpLUSTSite.ResumeLayout(False)
        Me.grpActiveTanks.ResumeLayout(False)
        Me.tbCtrlAdvancedSearch.ResumeLayout(False)
        Me.tbPageAdvancedSearch.ResumeLayout(False)
        Me.tbPageAdvancedSearchResult.ResumeLayout(False)
        Me.pnlSearchResult.ResumeLayout(False)
        CType(Me.ugAdvSearch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlSearchResultTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Advance Search Form Load"
    Private Sub AdvancedSearch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            oAdvanceSearch = New MUSTER.BusinessLogic.pAdvancedSearch
            frmRegServices = Me.MdiParent
            bolIsLoaded = False
            LoadFavSearches()
            bolIsLoaded = True

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#Region "FAV SEARCHES"
    Private Sub LoadFavSearches()
        Dim dtFavSearchList As DataTable
        Dim dRow As DataRow
        Try

            oAdvanceSearch.User = MusterContainer.AppUser.ID
            oAdvanceSearch.GetAll(oAdvanceSearch.User)
            dtFavSearchList = oAdvanceSearch.ParentTable()
            cmbFavoriteSearches.DataSource = Nothing
            cmbFavoriteSearches.DataSource = dtFavSearchList
            cmbFavoriteSearches.DisplayMember = "SEARCH_NAME"
            cmbFavoriteSearches.ValueMember = "SEARCH_ID"
            'If Not dtFavSearchList Is Nothing Then
            '    If dtFavSearchList.Rows.Count > 0 Then
            '        cmbFavoriteSearches.SelectedIndex = 0
            '    End If
            'End If
            cmbFavoriteSearches.SelectedIndex = -1

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#End Region
#Region "Advance Search Buttons Event Handlers"
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            bolAdvancedSearchError = False
            oAdvanceSearch.AddNewInfo(oAdvanceSearch.ID, Me.lstSearchBys.SelectedItem, Me.cmbSearchType.SelectedItem)
            If Not bolAdvancedSearchError Then
                BolNewSearch = True
                cmbFavoriteSearches_SelectedIndexChanged(cmbFavoriteSearches, e)
                If ugSearchByLookFor.Rows.Count > 0 Then ugSearchByLookFor.ActiveRow = ugSearchByLookFor.Rows(ugSearchByLookFor.Rows.Count - 1)
                BolNewSearch = False
            End If
        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Try
            If ugSearchByLookFor.Rows.Count <> 0 Then
                If Not ugSearchByLookFor.ActiveRow.Cells("CRITERION_ID").Value Is Nothing Then
                    oAdvanceSearch.RemoveChild(ugSearchByLookFor.ActiveRow.Cells("CRITERION_ID").Value)
                    oAdvanceSearch.GetParentByID(oAdvanceSearch.ID)
                    Me.FillGrid(oAdvanceSearch.ChildTable)
                Else
                    MessageBox.Show("Please Select a SearchBy Row to remove from the Collection")
                End If
            Else
                MessageBox.Show("Please Select a SearchBy Row to remove from the Collection")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Enum EnumLustStatus
        Open = 624
        Closed = 625
        OpenOrClosed = -1
        Unknown = 0
    End Enum
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click

        UCAdvSearchSummary.Hide()
        ugAdvSearch.Show()
        lstViewResults.Show()

        'For Adding Active Tanks option
        Dim dtDataTable As DataTable
        Dim nOwnerId As Integer
        Dim nFacId As Integer
        Dim dsResults As DataSet



        strTankStatus = oAdvanceSearch.TankStatus.ToUpper
        'If rdActiveUnknown.Checked = True Then
        '    strTankStatus = "UNKNOWN"
        'ElseIf rdActive.Checked = True Then
        '    strTankStatus = "ACTIVE"
        'ElseIf rdClose.Checked = True Then
        '    strTankStatus = "INACTIVE"
        'End If

        If oAdvanceSearch.LustStatus.ToUpper = "OPEN" Then
            nLustStatus = EnumLustStatus.Open
        ElseIf oAdvanceSearch.LustStatus.ToUpper = "CLOSED" Then
            nLustStatus = EnumLustStatus.Closed
        ElseIf oAdvanceSearch.LustStatus.ToUpper = "OPENORCLOSED" Then
            nLustStatus = EnumLustStatus.OpenOrClosed
        ElseIf oAdvanceSearch.LustStatus.ToUpper = "UNKNOWN" Then
            nLustStatus = EnumLustStatus.Unknown
        Else
            nLustStatus = EnumLustStatus.Unknown
        End If
        'If rdOpen.Checked Then
        '    nLustStatus = EnumLustStatus.Open
        'ElseIf rdClosed.Checked Then
        '    nLustStatus = EnumLustStatus.Closed
        'ElseIf rdLUSTOpenorClosed.Checked Then
        '    nLustStatus = EnumLustStatus.OpenOrClosed
        'ElseIf rdLUSTUnknown.Checked Then
        '    nLustStatus = EnumLustStatus.Unknown
        'End If

        If ugSearchByLookFor.Rows.Count = 0 Then
            MsgBox("You must supply at least one search criterion!", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "No Search Criteria Specified!")
            Exit Sub
        End If

   
        Try
            strModule = String.Empty
            nEventID = 0
            nClosureID = 0

            Me.Cursor = Windows.Forms.Cursors.WaitCursor

            'code for Search ALL option.
            If cmbSearchType.Text = "All" Then
                btnPrint.Enabled = False
                ugAdvSearch.Hide()
                lstViewResults.Hide()
                Me.pnlSearchResult.Controls.Add(UCAdvSearchSummary)
                UCAdvSearchSummary.Show()
                UCAdvSearchSummary.Dock = DockStyle.Fill
                tbCtrlAdvancedSearch.SelectedTab = tbPageAdvancedSearchResult
                UCAdvSearchSummary.pnlOwnerDetails.Visible = True
                UCAdvSearchSummary.pnlFacilitiesDetails.Visible = True
                UCAdvSearchSummary.pnlContactdetails.Visible = True
                UCAdvSearchSummary.pnlCompanydetails.Visible = True
                UCAdvSearchSummary.pnlContractorDetails.Visible = True
                UCAdvSearchSummary.lblOwnerDisplay.Text = "-"
                UCAdvSearchSummary.lblFacilitiesDisplay.Text = "-"
                UCAdvSearchSummary.lblContactsDisplay.Text = "-"
                UCAdvSearchSummary.lblCompanyDisplay.Text = "-"
                UCAdvSearchSummary.lblContractorDisplay.Text = "-"
                lblRecordCount.Text = ""
                lblSearchResult.Text = "Search Summary"
                SummaryThread = New Thread(AddressOf ProcessSearchSummary)
                SummaryThread.Start()
                Exit Sub
            End If
            btnPrint.Enabled = True

            If oAdvanceSearch.ID < 0 Then
                dsResults = oAdvanceSearch.GetResults(oAdvanceSearch.ID, cmbSearchType.SelectedItem, strTankStatus, nLustStatus)
            Else
                dsResults = oAdvanceSearch.GetResults(cmbFavoriteSearches.SelectedValue, cmbSearchType.SelectedItem, strTankStatus, nLustStatus)
            End If
            If Not dsResults Is Nothing Then
                If dsResults.Tables.Count > 0 Then
                    dtDataTable = dsResults.Tables(0)
                    dtDataTable = oAdvanceSearch.GetAdvSearchTable(cmbSearchType.SelectedItem, dtDataTable)
                    'If oAdvanceSearch.colIsDirty And oAdvanceSearch.ID > 0 Then  'Only when existing Favorite search.
                    '    Dim msgResult = MsgBox("WARNING!" & vbCrLf & "Your Favorite Search criteria has changed.you want to save it." & _
                    '                            " before executing the search?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo, "Data Modified")
                    '    If msgResult = MsgBoxResult.Yes Then
                    '        Me.Cursor = Windows.Forms.Cursors.WaitCursor
                    '        oAdvanceSearch.Flush()
                    '    End If
                    'End If
                    Select Case dtDataTable.Rows.Count
                        Case 0
                            RaiseEvent SearchResultSelection(0, 0, Nothing)
                            'Me.Cursor = Windows.Forms.Cursors.Default
                            Exit Sub
                        Case 1
                            If cmbSearchType.Text = "Owner" Then
                                nOwnerId = dtDataTable.Rows(0).Item("OWNER ID")
                                nFacId = 0
                            ElseIf cmbSearchType.Text = "Facility" Then
                                nOwnerId = dtDataTable.Rows(0).Item("OWNER ID")
                                nFacId = dtDataTable.Rows(0).Item("FACILITY ID")
                            ElseIf cmbSearchType.Text = "Contact" Then
                                If dtDataTable.Rows(0).Item("Contact Source") = "Lust Event" Then
                                    nOwnerId = dtDataTable.Rows(0).Item("Owner ID")
                                    nFacId = dtDataTable.Rows(0).Item("Facility ID")
                                    strModule = dtDataTable.Rows(0).Item("Module ID")
                                    nEventID = Integer.Parse(dtDataTable.Rows(0).Item("Contact Source ID"))
                                ElseIf dtDataTable.Rows(0).Item("Contact Source") = "Closure Event" Then
                                    nOwnerId = dtDataTable.Rows(0).Item("Owner ID")
                                    nFacId = dtDataTable.Rows(0).Item("Facility ID")
                                    strModule = dtDataTable.Rows(0).Item("Module ID")
                                    nClosureID = Integer.Parse(dtDataTable.Rows(0).Item("Contact Source ID"))
                                ElseIf dtDataTable.Rows(0).Item("Contact Source") = "Owner" Then
                                    nOwnerId = Integer.Parse(dtDataTable.Rows(0).Item("Contact Source ID"))
                                    nFacId = 0
                                    strModule = dtDataTable.Rows(0).Item("Module ID")
                                ElseIf dtDataTable.Rows(0).Item("Contact Source") = "Facility" Then
                                    nOwnerId = dtDataTable.Rows(0).Item("Owner ID")
                                    nFacId = Integer.Parse(dtDataTable.Rows(0).Item("Contact Source ID"))
                                    strModule = dtDataTable.Rows(0).Item("Module ID")
                                End If
                            ElseIf cmbSearchType.Text = "Company" Then
                                RaiseEvent CompanySearchSelection(dtDataTable.Rows(0).Item("Company ID"), 0, 0, "Company Name")
                                Exit Sub
                            ElseIf cmbSearchType.Text = "Contractor" Then
                                RaiseEvent CompanySearchSelection(IIf(dtDataTable.Rows(0).Item("Company ID") Is DBNull.Value, 0, dtDataTable.Rows(0).Item("Company ID")), dtDataTable.Rows(0).Item("Licensee ID"), 0, "Licensee Name")
                                Exit Sub
                            End If

                            RaiseEvent SearchResultSelection(nOwnerId, nFacId, cmbSearchType.Text)
                            ' Me.Cursor = Windows.Forms.Cursors.Default
                            Exit Sub
                        Case Else
                            lblRecordCount.Text = dtDataTable.Rows.Count.ToString() + "  Records Found"
                            lblSearchCriterion.Text = "LustStatus: " + oAdvanceSearch.LustStatus + ", TankStatus: " + oAdvanceSearch.TankStatus
                            Dim dtChildTable As DataTable = oAdvanceSearch.ChildTable
                            Dim dv As DataView = dtChildTable.DefaultView
                            dv.Sort = "CRITERION_ORDER"
                            For rowIndex As Integer = 0 To dtChildTable.Rows.Count - 1
                                lblSearchCriterion.Text += ", " + dv.Item(rowIndex)("CRITERION_NAME") + ": " + dv.Item(rowIndex)("CRITERION_VALUE")
                            Next
                            'lblRecordCount.Text = dtDataTable.Rows.Count.ToString() + "  Records Found."
                            tbCtrlAdvancedSearch.SelectedTab = tbPageAdvancedSearchResult
                            ugAdvSearch.DataSource = Nothing
                            ugAdvSearch.DataBind()
                            ugAdvSearch.DataSource = dtDataTable
                            ugAdvSearch.DataBind()
                            ugAdvSearch.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
                            ugAdvSearch.DisplayLayout.Bands(0).AutoPreviewEnabled = True
                            ugAdvSearch.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                            ugAdvSearch.Focus()
                            If cmbSearchType.Text = "Owner" Then
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("SNo").Width = 50
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner ID").Width = 75
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner Name").Width = 200
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner Address").Width = 250
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner Address").VertScrollBar = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner City").Width = 100
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("State").Width = 40
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner Contact").Width = 120
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner Phone").Width = 90
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner Points").Width = 75
                                lblSearchResult.Text = "Owner Search Results"
                            ElseIf cmbSearchType.Text = "Facility" Then
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("SNo").Width = 50
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility ID").Width = 75
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner Name").Width = 150
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility Name").Width = 200
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility Address").Width = 200
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility Address").VertScrollBar = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility City").Width = 100
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility County").Width = 80
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Latitude").Width = 165
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Longitude").Width = 165
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Points").Width = 50
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Lust Site").Width = 35
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("CIU").Width = 35
                                'ugAdvSearch.DisplayLayout.Bands(0).Columns("Total Tanks").Width = 50
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("TOS").Width = 35
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("POU").Width = 35
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("TOSI").Width = 35
                                'ugAdvSearch.DisplayLayout.Bands(0).Columns("Permanent Closure Pending").Width = 25
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility Contact").Width = 120
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility Phone").Width = 100
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner ID").Width = 75
                                lblSearchResult.Text = "Facility Search Results"
                            ElseIf cmbSearchType.Text = "Contact" Then
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Name").Width = 150
                                'ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Last Name").Width = 150
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Address").Width = 350
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Address").VertScrollBar = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                                'ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact City").Width = 100
                                'ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact State").Width = 80
                                'ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Phone").Width = 80
                                'ugAdvSearch.DisplayLayout.Bands(0).Columns("ZipCode").Width = 75
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Type").Width = 150
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Associated AT").Width = 200

                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Source").Width = 125
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Source").Hidden = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Source ID").Width = 100
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Source ID").Hidden = True

                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Contact Points").Width = 100
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Module ID").Hidden = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Owner ID").Hidden = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Facility ID").Hidden = True
                                lblSearchResult.Text = "Contact Search Results"
                            ElseIf cmbSearchType.Text = "Company" Then
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Company Name").Width = 180
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Company Address").Width = 200
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Company Address").VertScrollBar = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Company Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("City").Width = 100
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("State").Width = 40
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Phone").Width = 90
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Zip").Width = 45
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Installer/Closures").Width = 100
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Closures").Width = 65
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("ERAC").Width = 50
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Points").Width = 40
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Company ID").Hidden = True
                                lblSearchResult.Text = "Company Search Results"
                            ElseIf cmbSearchType.Text = "Contractor" Then
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Licensee Name").Width = 150
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Company Name").Width = 150
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Address").Width = 160
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Address").VertScrollBar = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("City").Width = 85
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("State").Width = 40
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Phone").Width = 85
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Zip").Width = 45
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Status").Width = 125
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Expiration Date").Width = 85
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Points").Width = 40
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Licensee ID").Hidden = True
                                ugAdvSearch.DisplayLayout.Bands(0).Columns("Company ID").Hidden = True
                                lblSearchResult.Text = "Contractor Search Results"
                            End If

                            If ugAdvSearch.Rows.Count > 0 Then
                                ugAdvSearch.ActiveRow = ugAdvSearch.Rows(0)
                            End If
                            Me.ugAdvSearch.DisplayLayout.Bands(0).Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                            ' Me.Cursor = Windows.Forms.Cursors.Default
                    End Select
                End If
            Else
                RaiseEvent SearchResultSelection(0, 0, Nothing)
                '  Me.Cursor = Windows.Forms.Cursors.Default
                Exit Sub
            End If



        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            dTableAdvSearch.Rows.Clear()
            bolIsLoaded = False
            If Not bolSearchTypeSelected = True Then
                cmbSearchType.SelectedIndex = -1
            End If
            If Not bolFavSearchSelected Then
                cmbFavoriteSearches.SelectedIndex = -1
                cmbFavoriteSearches.SelectedIndex = -1
            End If
            nSelIndexValue = cmbFavoriteSearches.SelectedIndex
            bolIsLoaded = True

            'cmbFavoriteSearches.SelectedIndex = -1
            lstViewResults.Items.Clear()
            If Not ugSearchByLookFor.DataSource Is Nothing Then
                'CType(ugSearchByLookFor.DataSource, DataTable).Clear()
                ugSearchByLookFor.DataSource = Nothing
            End If
            bolSearchTypeSelected = False
            bolFavSearchSelected = False
            oAdvanceSearch = New MUSTER.BusinessLogic.pAdvancedSearch
            oAdvanceSearch.GetAll(MusterContainer.AppUser.ID)
            LoadChkBoxes()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        'If oAdvanceSearch.colIsDirty Then
        '    If (MsgBox("Do you want to save the changes made to your searches?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
        '        oAdvanceSearch.Flush()
        '    End If
        'End If
        Me.Dispose()

    End Sub
    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Try
            Dim OldRow As Int16 = ugSearchByLookFor.ActiveRow.Index
            bolAdvancedSearchError = False
            oAdvanceSearch.ChangeOrder(ugSearchByLookFor.ActiveRow.Cells("CRITERION_ORDER").Value, dtFavSearchChildList, -1)
            If Not bolAdvancedSearchError Then
                FillGrid(dtFavSearchChildList)
                ugSearchByLookFor.ActiveRow = ugSearchByLookFor.Rows(OldRow - 1)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click
        Try
            bolAdvancedSearchError = False
            Dim OldRow As Int16 = ugSearchByLookFor.ActiveRow.Index
            oAdvanceSearch.ChangeOrder(ugSearchByLookFor.ActiveRow.Cells("CRITERION_ORDER").Value, dtFavSearchChildList, 1)
            If Not bolAdvancedSearchError Then
                FillGrid(dtFavSearchChildList)
                ugSearchByLookFor.ActiveRow = ugSearchByLookFor.Rows(OldRow + 1)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSaveSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveSearch.Click
        Try
            If Me.cmbSearchType.SelectedIndex <> -1 Then
                If ugSearchByLookFor.Rows.Count = 0 Then
                    MsgBox("You must supply at least one search criterion!", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "No Search Criteria Specified!")
                    Exit Sub
                End If
                frmConfirmSave = New FavSearchDialog(oAdvanceSearch)
                frmConfirmSave.frmParent = Me
                frmConfirmSave.ShowInTaskbar = False
                If oAdvanceSearch.ID > 0 Then
                    frmConfirmSave.txtFav_Search_Name.Text = cmbFavoriteSearches.Text
                End If
                frmConfirmSave.ShowDialog()
            Else
                MsgBox("Please select SearchType.!", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "No SearchType Specified!")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MsgBox("Are you sure you want to Delete this search?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                If cmbFavoriteSearches.SelectedIndex <> -1 Then
                    oAdvanceSearch.RemoveParent(cmbFavoriteSearches.SelectedValue)
                    bolIsLoaded = False
                    LoadFavSearches()
                    bolIsLoaded = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub tbCtrlAdvancedSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlAdvancedSearch.Click
        Select Case tbCtrlAdvancedSearch.SelectedTab.Name.ToUpper
            Case "TBPAGEADVANCEDSEARCH"
                bolTabClickEvent = True
                If IsNumeric(cmbFavoriteSearches.Tag) Then
                    cmbFavoriteSearches.SelectedIndex = cmbFavoriteSearches.Tag
                    If Integer.Parse(cmbFavoriteSearches.Tag) = -1 Then cmbFavoriteSearches.SelectedIndex = cmbFavoriteSearches.Tag
                End If
                bolTabClickEvent = False
            Case "TBPAGEADVANCEDSEARCHRESULT"
                cmbFavoriteSearches.Tag = cmbFavoriteSearches.SelectedIndex
        End Select
    End Sub
    Private Sub AdvancedSearch_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        ugAdvSearch.PrintPreview(Infragistics.Win.UltraWinGrid.RowPropertyCategories.All)
    End Sub
#End Region
#Region "Combo Box Change Events"
    Private Sub cmbSearchType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSearchType.SelectedIndexChanged
        Try
            If cmbSearchType.SelectedIndex = -1 Or Not bolIsLoaded Then
                Exit Sub
            End If
            If bolTabClickEvent Then Exit Sub


            bolSearchTypeSelected = True
            btnClear_Click(btnClear, New System.EventArgs)

            'ugSearchByLookFor.DataSource = Nothing
            'We have to add the parent search Id with -1 and child with -1 AND CHILD NAME DEFUALT VALUES
            oAdvanceSearch.ResetCollection()
            oAdvanceSearch.Search_Type = Me.cmbSearchType.SelectedItem
            oAdvanceSearch.AddNewInfo(oAdvanceSearch.ID, cmbSearchType.Text, Me.cmbSearchType.SelectedItem)

            If rdOpen.Checked Then
                oAdvanceSearch.LustStatus = "Open"
            ElseIf rdClosed.Checked Then
                oAdvanceSearch.LustStatus = "Closed"
            ElseIf rdLUSTOpenorClosed.Checked Then
                oAdvanceSearch.LustStatus = "OpenOrClosed"
            ElseIf rdLUSTUnknown.Checked Then
                oAdvanceSearch.LustStatus = "Unknown"
            Else
                oAdvanceSearch.LustStatus = "Unknown"
            End If

            If rdActiveUnknown.Checked = True Then
                oAdvanceSearch.TankStatus = "Unknown"
            ElseIf rdActive.Checked = True Then
                oAdvanceSearch.TankStatus = "Active"
            ElseIf rdClose.Checked = True Then
                oAdvanceSearch.TankStatus = "InActive"
            Else
                oAdvanceSearch.TankStatus = "InActive"
            End If

            BolNewSearch = True
            cmbFavoriteSearches_SelectedIndexChanged(sender, e)
            BolNewSearch = False
            If cmbSearchType.SelectedIndex <> 2 Then
                frmRegServices.cmbSearchModule.SelectedValue = MusterContainer.AppUser.DefaultModule
            End If
            Dim item As ListViewItem
            Dim dr As DataRow
            Dim dtFilterList As DataTable
            dtFilterList = GetFilterList(cmbSearchType.Text)
            lstSearchBys.Items.Clear()
            For Each dr In dtFilterList.Rows
                lstSearchBys.Items.Add(dr.Item("FilterName"))
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cmbFavoriteSearches_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFavoriteSearches.SelectedIndexChanged

        Try
            If Not bolIsLoaded Then Exit Sub
            If bolTabClickEvent Then Exit Sub
            nSelIndexValue = cmbFavoriteSearches.SelectedIndex
            'If CType(sender, System.Windows.Forms.ComboBox).SelectedIndex = -1 Then Exit Sub

            Dim dtCriteria As DataTable = CType(ugSearchByLookFor.DataSource, DataTable)
            If Not dtCriteria Is Nothing Then
                dtCriteria.Rows.Clear()
            End If

            If Not cmbFavoriteSearches.SelectedValue < 0 And Not cmbFavoriteSearches.SelectedValue = Nothing And Not BolNewSearch Then
                bolFavSearchSelected = True
                btnClear_Click(btnClear, New System.EventArgs)
            End If
            If BolNewSearch Then 'New 
                dtCriteria = oAdvanceSearch.GetCriteria(oAdvanceSearch.ID)
            Else
                If cmbFavoriteSearches.SelectedIndex < 0 Then Exit Sub
                dtCriteria = oAdvanceSearch.GetCriteria(cmbFavoriteSearches.SelectedValue)
                bolIsLoaded = False
                cmbSearchType.SelectedItem = oAdvanceSearch.SearchType
                bolIsLoaded = True
            End If
            nSelIndexValue = cmbFavoriteSearches.SelectedIndex
            cmbFavoriteSearches.Tag = cmbFavoriteSearches.SelectedIndex
            FillGrid(dtCriteria)
            LoadChkBoxes()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FillGrid(ByVal dtFavSearch As DataTable)

        Try
            dtFavSearchChildList = dtFavSearch
            dtFavSearchChildList.DefaultView.Sort = "CRITERION_ORDER ASC"
            ugSearchByLookFor.DataSource = dtFavSearchChildList
            ugSearchByLookFor.DataBind()
            ugSearchByLookFor.DisplayLayout.Bands(0).Columns("SEARCH_ID").Hidden = True
            ugSearchByLookFor.DisplayLayout.Bands(0).Columns("CRITERION_DATA_TYPE").Hidden = True
            ugSearchByLookFor.DisplayLayout.Bands(0).Columns("CRITERION_ID").Hidden = True
            ugSearchByLookFor.DisplayLayout.Bands(0).Columns("CRITERION_ORDER").Hidden = True
            ugSearchByLookFor.DisplayLayout.Appearance.ForeColorDisabled = ugSearchByLookFor.DisplayLayout.Appearance.ForeColor
            ugSearchByLookFor.DisplayLayout.Bands(0).Columns("CRITERION_NAME").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            ugSearchByLookFor.DisplayLayout.Bands(0).Columns("CRITERION_VALUE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadChkBoxes()
        Try
            If oAdvanceSearch.LustStatus.ToUpper = "OPEN" Then
                rdOpen.Checked = True
            ElseIf oAdvanceSearch.LustStatus.ToUpper = "CLOSED" Then
                rdClosed.Checked = True
            ElseIf oAdvanceSearch.LustStatus.ToUpper = "OPENORCLOSED" Then
                rdLUSTOpenorClosed.Checked = True
            ElseIf oAdvanceSearch.LustStatus.ToUpper = "UNKNOWN" Then
                rdLUSTUnknown.Checked = True
            Else
                rdLUSTUnknown.Checked = True
            End If

            If oAdvanceSearch.TankStatus.ToUpper = "INACTIVE" Then
                rdClose.Checked = True
            ElseIf oAdvanceSearch.TankStatus.ToUpper = "ACTIVE" Then
                rdActive.Checked = True
            ElseIf oAdvanceSearch.TankStatus.ToUpper = "UNKNOWN" Then
                rdActiveUnknown.Checked = True
            Else
                rdActiveUnknown.Checked = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "ChkBox Change Events"
    Private Sub rdOpen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdOpen.CheckedChanged
        If oAdvanceSearch Is Nothing Then Exit Sub
        If rdOpen.Checked Then
            oAdvanceSearch.LustStatus = "Open"
        End If
    End Sub
    Private Sub rdClosed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdClosed.CheckedChanged
        If oAdvanceSearch Is Nothing Then Exit Sub
        If rdClosed.Checked Then
            oAdvanceSearch.LustStatus = "Closed"
        End If
    End Sub
    Private Sub rdLUSTOpenorClosed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdLUSTOpenorClosed.CheckedChanged
        If oAdvanceSearch Is Nothing Then Exit Sub
        If rdLUSTOpenorClosed.Checked Then
            oAdvanceSearch.LustStatus = "OpenOrClosed"
        End If
    End Sub
    Private Sub rdLUSTUnknown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdLUSTUnknown.CheckedChanged
        If oAdvanceSearch Is Nothing Then Exit Sub
        If rdLUSTUnknown.Checked Then
            oAdvanceSearch.LustStatus = "Unknown"
        End If
    End Sub
    Private Sub rdActiveUnknown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdActiveUnknown.CheckedChanged
        If oAdvanceSearch Is Nothing Then Exit Sub
        If rdActiveUnknown.Checked Then
            oAdvanceSearch.TankStatus = "Unknown"
        End If
    End Sub
    Private Sub rdActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdActive.CheckedChanged
        If oAdvanceSearch Is Nothing Then Exit Sub
        If rdActive.Checked Then
            oAdvanceSearch.TankStatus = "Active"
        End If
    End Sub
    Private Sub rdClose_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdClose.CheckedChanged
        If oAdvanceSearch Is Nothing Then Exit Sub
        If rdClose.Checked Then
            oAdvanceSearch.TankStatus = "InActive"
        End If
    End Sub
#End Region
#Region "Ultra Grid Events"
    Private Sub ugSearchByLookFor_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugSearchByLookFor.AfterCellUpdate
        Try
            If e.Cell Is Nothing Then
                Exit Sub
            End If
            Dim header As String
            header = e.Cell.Column.Header.Caption
            Select Case header
                Case "CRITERION_NAME"
                    oAdvanceSearch.ChildName = e.Cell.Text
                Case "CRITERION_VALUE"
                    oAdvanceSearch.CriterionValue = e.Cell.Text
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugAdvSearch_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugAdvSearch.AfterSortChange
        ugAdvSearch.ActiveRow = ugAdvSearch.Rows(0)
    End Sub
    Private Sub ugAdvSearch_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAdvSearch.DoubleClick
        Try
            Me.Cursor = Windows.Forms.Cursors.WaitCursor
            Dim nFacId As Integer
            Dim nOwnerId As Integer
            strModule = String.Empty
            nEventID = 0
            nClosureID = 0
            'MusterContainer.bolSearchModuleSelected = False
            If UIUtilsInfragistics.WinGridRowDblClicked(ugAdvSearch, New System.EventArgs) Then

                'RaiseEvent SearchResultSelection(ugAdvSearch.ActiveRow.Cells("Owner ID").Value, ugAdvSearch.ActiveRow.Cells("Facility ID").Value, cmbSearchType.Text)
                If cmbSearchType.Text = "Owner" Then
                    nOwnerId = ugAdvSearch.ActiveRow.Cells("Owner ID").Value
                    nFacId = 0
                ElseIf cmbSearchType.Text = "Facility" Then
                    nOwnerId = ugAdvSearch.ActiveRow.Cells("Owner ID").Value
                    nFacId = ugAdvSearch.ActiveRow.Cells("Facility ID").Value
                ElseIf cmbSearchType.Text = "Contact" Then
                    Dim nEntityTypeID As Integer = 0

                    If ugAdvSearch.ActiveRow.Cells("Contact Source").Value = "Lust Event" Then
                        nOwnerId = ugAdvSearch.ActiveRow.Cells("Owner ID").Value
                        nFacId = ugAdvSearch.ActiveRow.Cells("Facility ID").Value
                        strModule = ugAdvSearch.ActiveRow.Cells("Module ID").Value
                        nEventID = ugAdvSearch.ActiveRow.Cells("Contact Source ID").Value
                        nEntityTypeID = UIUtilsGen.EntityTypes.LUST_Event
                    ElseIf ugAdvSearch.ActiveRow.Cells("Contact Source").Value = "Closure Event" Then
                        nOwnerId = ugAdvSearch.ActiveRow.Cells("Owner ID").Value
                        nFacId = ugAdvSearch.ActiveRow.Cells("Facility ID").Value
                        strModule = ugAdvSearch.ActiveRow.Cells("Module ID").Value
                        nClosureID = Integer.Parse(ugAdvSearch.ActiveRow.Cells("Contact Source ID").Value)
                        nEntityTypeID = UIUtilsGen.EntityTypes.ClosureEvent
                    ElseIf ugAdvSearch.ActiveRow.Cells("Contact Source").Value = "Owner" Then
                        nOwnerId = Integer.Parse(Integer.Parse(ugAdvSearch.ActiveRow.Cells("Contact Source ID").Value))
                        nFacId = 0
                        strModule = ugAdvSearch.ActiveRow.Cells("Module ID").Value
                        nEntityTypeID = UIUtilsGen.EntityTypes.Owner
                    ElseIf ugAdvSearch.ActiveRow.Cells("Contact Source").Value = "Facility" Then
                        nOwnerId = ugAdvSearch.ActiveRow.Cells("Owner ID").Value
                        nFacId = Integer.Parse(Integer.Parse(ugAdvSearch.ActiveRow.Cells("Contact Source ID").Value))
                        strModule = ugAdvSearch.ActiveRow.Cells("Module ID").Value
                        nEntityTypeID = UIUtilsGen.EntityTypes.Facility
                    ElseIf ugAdvSearch.ActiveRow.Cells("Contact Source").Value = "FinancialEvent" Then
                        nOwnerId = ugAdvSearch.ActiveRow.Cells("Owner ID").Value
                        nFacId = ugAdvSearch.ActiveRow.Cells("Facility ID").Value
                        strModule = ugAdvSearch.ActiveRow.Cells("Module ID").Value
                        nEventID = ugAdvSearch.ActiveRow.Cells("Contact Source ID").Value
                        nEntityTypeID = UIUtilsGen.EntityTypes.FinancialEvent
                    End If

                    Dim moduleID As Integer = UIUtilsGen.GetModuleIDByName(ugAdvSearch.ActiveRow.Cells("Module ID").Value)
                    If Not MusterContainer.AppUser.HasAccess(moduleID, MusterContainer.AppUser.UserKey, nEntityTypeID) Then
                        MessageBox.Show("You do not have rights to view " + ugAdvSearch.ActiveRow.Cells("Contact Source").Value.ToString)
                        Exit Sub
                    End If


                ElseIf cmbSearchType.Text = "Company" Then
                    RaiseEvent CompanySearchSelection(ugAdvSearch.ActiveRow.Cells("Company ID").Value, 0, 0, "Company Name")
                    Exit Sub
                ElseIf cmbSearchType.Text = "Contractor" Then
                    RaiseEvent CompanySearchSelection(IIf(ugAdvSearch.ActiveRow.Cells("Company ID").Value Is System.DBNull.Value, 0, ugAdvSearch.ActiveRow.Cells("Company ID").Value), ugAdvSearch.ActiveRow.Cells("Licensee ID").Value, 0, "Licensee Name")
                    Exit Sub
                End If
                RaiseEvent SearchResultSelection(nOwnerId, nFacId, cmbSearchType.Text)

            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Windows.Forms.Cursors.Default
        End Try
    End Sub
    Private Sub ugAdvSearch_InitializePrintPreview(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelablePrintPreviewEventArgs) Handles ugAdvSearch.InitializePrintPreview
        UIUtilsGen.ug_InitializePrintPreview(sender, e, oAdvanceSearch.SearchType + " Search - " + lblSearchCriterion.Text, _
                                        Infragistics.Win.UltraWinGrid.ColumnClipMode.Default, _
                                        False, Printing.PrintRange.AllPages, _
                                        True, , Infragistics.Win.DefaultableBoolean.True, , 1)
    End Sub
    'Private Sub ugSearchByLookFor_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugSearchByLookFor.AfterRowActivate
    '    Try
    '        If Not IsDBNull(ugSearchByLookFor.ActiveRow.Cells(0).Value) Then
    '            oAdvanceSearch.GetChildByID(ugSearchByLookFor.ActiveRow.Cells(0).Value)
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub ugSearchByLookFor_BeforeRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugSearchByLookFor.BeforeRowActivate
        Try
            If e.Row.Cells.Count > 0 Then
                If Not e.Row.Cells(0).Value Is DBNull.Value Then
                    oAdvanceSearch.GetChildByID(e.Row.Cells(0).Value)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "External Events"
    Private Sub frmConfirmSave_SaveFavoriteSearch(ByVal strFavSearchName As String) Handles frmConfirmSave.SaveFavoriteSearch
        Try
            'oAdvanceSearch.Flush()
            oAdvanceSearch.SaveParent()
            bolIsLoaded = False
            LoadFavSearches()
            bolIsLoaded = True
            cmbFavoriteSearches.Text = strFavSearchName
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub AdvancedSearchError(ByVal strMessage As String, ByVal strSRC As String) Handles oAdvanceSearch.AdvancedSearchErr
        bolAdvancedSearchError = True
        MessageBox.Show(strMessage)
    End Sub
    Private Sub EnableDisableDelete(ByVal nCount As Integer, ByVal strSRC As String) Handles oAdvanceSearch.eEnableDisableDelete
        If nCount = 0 Then
            btnDelete.Enabled = False
        Else
            btnDelete.Enabled = True
        End If
    End Sub
#End Region
#Region "Summary Search Procedures"

    Private Sub ProcessSearchSummary()
        Invoke(New OwnerSearchHandler(AddressOf OwnerSummary))
        Invoke(New OwnerSearchHandler(AddressOf FacilitySummary))
        Invoke(New OwnerSearchHandler(AddressOf ContactSummary))
        Invoke(New OwnerSearchHandler(AddressOf CompanySummary))
        Invoke(New OwnerSearchHandler(AddressOf ContractorSummary))
    End Sub
    Private Sub OwnerSummary()

        Dim dsSummaryResult As DataSet
        Dim SelectedSearchType As String = String.Empty
        Dim dtTable As DataTable
        'Owner
        SelectedSearchType = "Owner"
        If oAdvanceSearch.ID < 0 Then
            dsSummaryResult = oAdvanceSearch.GetResults(oAdvanceSearch.ID, SelectedSearchType, strTankStatus, nLustStatus)
        Else
            dsSummaryResult = oAdvanceSearch.GetResults(cmbFavoriteSearches.SelectedValue, SelectedSearchType, strTankStatus, nLustStatus)
        End If

        If dsSummaryResult.Tables.Count > 0 Then
            dtTable = dsSummaryResult.Tables(0)
            dtTable = oAdvanceSearch.GetAdvSearchTable(SelectedSearchType, dtTable)

            UCAdvSearchSummary.ugOwner.DataSource = Nothing
            UCAdvSearchSummary.ugOwner.DataBind()
            UCAdvSearchSummary.ugOwner.DataSource = dtTable
            UCAdvSearchSummary.ugOwner.DataBind()
            UCAdvSearchSummary.ugOwner.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            UCAdvSearchSummary.ugOwner.Focus()
            UCAdvSearchSummary.lblOwnerHeader.Text = "Owner             " + dtTable.Rows.Count.ToString() + " Records Found"
        End If

    End Sub
    Private Sub FacilitySummary()

        Dim dsSummaryResult As DataSet
        Dim SelectedSearchType As String = String.Empty
        Dim dtTable As DataTable
        'Facility
        SelectedSearchType = "Facility"
        If oAdvanceSearch.ID < 0 Then
            dsSummaryResult = oAdvanceSearch.GetResults(oAdvanceSearch.ID, SelectedSearchType, strTankStatus, nLustStatus)
        Else
            dsSummaryResult = oAdvanceSearch.GetResults(cmbFavoriteSearches.SelectedValue, SelectedSearchType, strTankStatus, nLustStatus)
        End If

        If dsSummaryResult.Tables.Count > 0 Then
            dtTable = dsSummaryResult.Tables(0)
            dtTable = oAdvanceSearch.GetAdvSearchTable(SelectedSearchType, dtTable)

            UCAdvSearchSummary.ugFacilities.DataSource = Nothing
            UCAdvSearchSummary.ugFacilities.DataBind()
            UCAdvSearchSummary.ugFacilities.DataSource = dtTable
            UCAdvSearchSummary.ugFacilities.DataBind()
            UCAdvSearchSummary.ugFacilities.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            UCAdvSearchSummary.lblFacilitiesHeader.Text = "Facilities             " + dtTable.Rows.Count.ToString() + " Records Found"
        End If

    End Sub
    Private Sub ContactSummary()

        Dim dsSummaryResult As DataSet
        Dim SelectedSearchType As String = String.Empty
        Dim dtTable As DataTable
        'Contact
        SelectedSearchType = "Contact"
        If oAdvanceSearch.ID < 0 Then
            dsSummaryResult = oAdvanceSearch.GetResults(oAdvanceSearch.ID, SelectedSearchType, strTankStatus, nLustStatus)
        Else
            dsSummaryResult = oAdvanceSearch.GetResults(cmbFavoriteSearches.SelectedValue, SelectedSearchType, strTankStatus, nLustStatus)
        End If

        If dsSummaryResult.Tables.Count > 0 Then
            dtTable = dsSummaryResult.Tables(0)
            dtTable = oAdvanceSearch.GetAdvSearchTable(SelectedSearchType, dtTable)

            UCAdvSearchSummary.ugContacts.DataSource = Nothing
            UCAdvSearchSummary.ugContacts.DataBind()
            UCAdvSearchSummary.ugContacts.DataSource = dtTable
            UCAdvSearchSummary.ugContacts.DataBind()
            UCAdvSearchSummary.ugContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            UCAdvSearchSummary.lblContactHeader.Text = "Contacts             " + dtTable.Rows.Count.ToString() + " Records Found"
        End If

    End Sub
    Private Sub CompanySummary()

        Dim dsSummaryResult As DataSet
        Dim SelectedSearchType As String = String.Empty
        Dim dtTable As DataTable
        'Company
        SelectedSearchType = "Company"
        If oAdvanceSearch.ID < 0 Then
            dsSummaryResult = oAdvanceSearch.GetResults(oAdvanceSearch.ID, SelectedSearchType, strTankStatus, nLustStatus)
        Else
            dsSummaryResult = oAdvanceSearch.GetResults(cmbFavoriteSearches.SelectedValue, SelectedSearchType, strTankStatus, nLustStatus)
        End If

        If dsSummaryResult.Tables.Count > 0 Then
            dtTable = dsSummaryResult.Tables(0)
            dtTable = oAdvanceSearch.GetAdvSearchTable(SelectedSearchType, dtTable)

            UCAdvSearchSummary.ugCompanies.DataSource = Nothing
            UCAdvSearchSummary.ugCompanies.DataBind()
            UCAdvSearchSummary.ugCompanies.DataSource = dtTable
            UCAdvSearchSummary.ugCompanies.DataBind()
            UCAdvSearchSummary.ugCompanies.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            UCAdvSearchSummary.lblCompanyHeader.Text = "Company             " + dtTable.Rows.Count.ToString() + " Records Found"
        End If

    End Sub
    Private Sub ContractorSummary()

        Dim dsSummaryResult As DataSet
        Dim SelectedSearchType As String = String.Empty
        Dim dtTable As DataTable
        'Contractor
        SelectedSearchType = "Contractor"
        If oAdvanceSearch.ID < 0 Then
            dsSummaryResult = oAdvanceSearch.GetResults(oAdvanceSearch.ID, SelectedSearchType, strTankStatus, nLustStatus)
        Else
            dsSummaryResult = oAdvanceSearch.GetResults(cmbFavoriteSearches.SelectedValue, SelectedSearchType, strTankStatus, nLustStatus)
        End If

        If dsSummaryResult.Tables.Count > 0 Then
            dtTable = dsSummaryResult.Tables(0)
            dtTable = oAdvanceSearch.GetAdvSearchTable(SelectedSearchType, dtTable)

            UCAdvSearchSummary.ugContractors.DataSource = Nothing
            UCAdvSearchSummary.ugContractors.DataBind()
            UCAdvSearchSummary.ugContractors.DataSource = dtTable
            UCAdvSearchSummary.ugContractors.DataBind()
            UCAdvSearchSummary.ugContractors.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            UCAdvSearchSummary.lblContractorHeader.Text = "Contractors             " + dtTable.Rows.Count.ToString() + " Records Found"
        End If

    End Sub

#End Region
    Private Sub ugSearchByLookFor_AfterCellActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugSearchByLookFor.AfterCellActivate
        Try
            If ugSearchByLookFor.ActiveRow Is Nothing Then Exit Sub
            If UIUtilsInfragistics.WinGridRowDblClicked(ugSearchByLookFor, New System.EventArgs) Then
                ugSearchByLookFor.ActiveRow.Selected = True
                nLastLookForRow = ugSearchByLookFor.ActiveRow.Index
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try


    End Sub
    Private Sub ugSearchByLookFor_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ugSearchByLookFor.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnSearch.Focus()
            btnSearch.PerformClick()
        Else
            If ugSearchByLookFor.ActiveRow.Index = ugSearchByLookFor.Rows.Count - 1 And e.KeyCode = Keys.Tab And Not e.Shift Then
                btnUp.Focus()
            Else
                If ugSearchByLookFor.ActiveRow.Index = 0 And e.KeyCode = Keys.Tab And e.Shift Then
                    btnRemove.Focus()
                End If
            End If
        End If
    End Sub
    Private Sub ugSearchByLookFor_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugSearchByLookFor.GotFocus
        If ugSearchByLookFor.Rows.Count > 0 Then
            ugSearchByLookFor.ActiveRow = ugSearchByLookFor.Rows(0)
        End If
    End Sub

    Private Sub cmbFavoriteSearches_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFavoriteSearches.LostFocus
        If CType(sender, Windows.Forms.ComboBox).Text.Trim = String.Empty Then
            btnClear_Click(btnClear, New System.EventArgs)
        End If
    End Sub
    Private Function GetFilterList(ByVal SearchType As String) As DataTable
        Dim dtFilterList As New DataTable
        Dim drFilter As DataRow

     
        If SearchType = "Company" Or SearchType = "Contractor" Then
            If Not dtFilterList Is Nothing Then
                dtFilterList.Clear()
                'dtFilterList = Nothing

            End If
            dtFilterList.Columns.Add("FilterName")
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Company Name"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Licensee Name"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Type of Service"
            dtFilterList.Rows.Add(drFilter)
            'drFilter = dtFilterList.NewRow
            'drFilter("FilterName") = "Licensee First Name"
            'dtFilterList.Rows.Add(drFilter)
            'drFilter = dtFilterList.NewRow
            'drFilter("FilterName") = "Licensee Middle Name"
            'dtFilterList.Rows.Add(drFilter)
            'drFilter = dtFilterList.NewRow
            'drFilter("FilterName") = "Licensee Last Name"
            'dtFilterList.Rows.Add(drFilter)
            grpActiveTanks.Enabled = False
            grpLUSTSite.Enabled = False
        ElseIf SearchType = "All" Then
            dtFilterList.Columns.Add("FilterName")
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Name"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Address"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "City"
            dtFilterList.Rows.Add(drFilter)
            grpActiveTanks.Enabled = False
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Phone"
            dtFilterList.Rows.Add(drFilter)
            grpActiveTanks.Enabled = True
            grpLUSTSite.Enabled = True
        Else
            If Not dtFilterList Is Nothing Then
                dtFilterList.Clear()
            End If

            dtFilterList.Columns.Add("FilterName")

            If SearchType = "Facility" Then
                drFilter = dtFilterList.NewRow
                drFilter("FilterName") = "Facility Lat Degree"
                dtFilterList.Rows.Add(drFilter)
                drFilter = dtFilterList.NewRow

                drFilter("FilterName") = "Facility Lat Minutes"
                dtFilterList.Rows.Add(drFilter)
                drFilter = dtFilterList.NewRow
                drFilter("FilterName") = "Facility Long Degree"
                dtFilterList.Rows.Add(drFilter)
                drFilter = dtFilterList.NewRow

                drFilter("FilterName") = "Facility Long Minutes"
                dtFilterList.Rows.Add(drFilter)


            End If

            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "All"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Brand"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Facility Address"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Facility City"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Facility County"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Facility Name"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Facility AIID"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Owner Address"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Owner City"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Owner ID"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Owner Name"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Project Manager"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Project Manager (History)"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Contact Company Name"
            dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "Contact Name"
            dtFilterList.Rows.Add(drFilter)
            'drFilter = dtFilterList.NewRow
            'drFilter("FilterName") = "Contact First Name"
            'dtFilterList.Rows.Add(drFilter)
            'drFilter = dtFilterList.NewRow
            'drFilter("FilterName") = "Contact Middle Name"
            'dtFilterList.Rows.Add(drFilter)
            'drFilter = dtFilterList.NewRow
            'drFilter("FilterName") = "Contact Last Name"
            'dtFilterList.Rows.Add(drFilter)
            drFilter = dtFilterList.NewRow
            drFilter("FilterName") = "All Phone Numbers"
            dtFilterList.Rows.Add(drFilter)
            grpActiveTanks.Enabled = True
            grpLUSTSite.Enabled = True
        End If
        Return dtFilterList
    End Function

End Class

