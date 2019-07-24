'-------------------------------------------------------------------------------
' MUSTER.MUSTER.MusterContainer.vb
'   Provides the parent MDI window for the application.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        ??      8/??/04    Original class definition.
'  1.1        JC      1/02/04    Changed AppUser.UserName to AppUser.ID to
'                                  accomodate new use of pUser by application.
'  1.2        JVC2    01/20/05   Added functionality to open Calendar form properly.          
'  1.3        JVC2    01/27/05   Added code to searchresultselected to transfer
'                                  control to the appropriate module.
'  1.4        MR      01/27/05   Modified and Added Calendar Functions to Implement Calendar Object.
'  1.5        MR      01/28/05   Added LoadCalendarforSelectedUser() Function for Loading ViewEntries.
'  1.6        JVC2    02/03/05   Added new code to hook in Advanced Search through
'                                  the QuickSearch Results handler.
'  1.7        AN      02/10/05    Integrated AppFlags new object model
'  1.8        MNR     03/08/05   Modified pOwn.Retrieve / pOwn.Facilities.Retrieve to pass true for bolLoading paramater
'  1.9        JVC     03/14/05   Added DomainUser as global shared object
'  2.0        PN      03/15/05   Updated pOwn.Retrieve with pOwn.RetrieveAll (Speed Loading)
'  2.1        JVC     08/02/05   Changed intialization to load cmbSearchModule from
'                                   property list for type Modules.  Also changed
'                                   selection of cmbSearchModule to reflect the user's
'                                   default module as specified in Admin.
'  2.2        JVC2    08/09/05   Added code to ColorInOverDueEntries() to color entries
'                                   made by users other than the logged in user aqua (see DDD p 72)
'                                   Added txtOwnerQSKeyword_KeyUp event to process ENTER key
'                                   properly for quick search feature.


'   2.3        Thomas Franey   02/19/2009    Added Invoice Request Admin menu item in Finance Admin  
'                                            ;line:1826,1949,1957
'                              02/20/2009    Added Last change to be Dialog Window instead of MDi child
'
'   2.35       Thomas Franey   02/23/2009     Added Removehandler to line 3744 ro retain only one instance of frm close 
'                                             handling (line 3707) 
'
'                                             Added Event Subroutine to attach the global sort memory to each form opened
'                                             Added the form activated handler line for sort subroutine in line 3747. 
'
'  3.0         Thomas Franey    5/1/2009     Added members and handlers for Tickler System
'              Hua Cao          09/12/12     Added "Add compliance Manager" menu
'-------------------------------------------------------------------------------
'
' TODO - Integrate with application 2/9/05 JVC2
'
Imports System.Security.Principal
Imports System.Configuration.Assemblies
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Diagnostics

Imports SQLDMO.RestoreClass



Public Class MusterContainer
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"

    'Ticker members
    Public WithEvents TicklerTimerThread As System.Timers.Timer
    Public WithEvents TicklerAutoOpenThread As System.Timers.Timer


    Public AutoOpen As Boolean = False
    Public DirtyIgnored As Long = -1

    Public GoToEntityCode As Integer = -1
    Public GoToGrid As String = String.Empty
    Public GoToUser As String = String.Empty
    Public GoToTabPage As String = String.Empty
    Public GotoYear As String = String.Empty
    Public GotoReport As String = String.Empty
    Public GotoModule As String = String.Empty
    Public TAlert As TicklerAlert




    Public Inspector As Boolean = False

    Dim CallOpenForm As New MethodInvoker(AddressOf InvokeQuickOwneerButtonClick)
    Dim CallDocumentManager As New MethodInvoker(AddressOf InvokeMnulettersClick)
    Dim CallReportManager As New MethodInvoker(AddressOf InvokeReportClick)


    'muster container members
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    Public Shared MdiChildClosing As Boolean = False
    Public ParentApp As System.Windows.Forms.Form
    'Public objFees As New MUSTER.FeesServices
    Public frmAdvanceSearch As AdvancedSearch = Nothing
    Public Shared AppUser As MUSTER.BusinessLogic.pUser
    Public Shared AppProfileInfo As MUSTER.BusinessLogic.pProfile
    Public Shared WithEvents AppSemaphores As MUSTER.BusinessLogic.pAppFlag
    Private strModuleTitle As String

    Public Shared ProfileData As MUSTER.BusinessLogic.pProfile
    Friend WithEvents objRegister As Registration
    Private WithEvents _thisForm As MusterContainer
    Friend Shared MyGUID As System.Guid
    Private bolLoadingForm As Boolean = False
    Friend Shared flag As Boolean

    Dim boolShowImages As Boolean
    Private bolErrorOccurred As Boolean = False
    Private WithEvents oQS As MUSTER.BusinessLogic.pSearch
    Public Shared WithEvents pCalendar As MUSTER.BusinessLogic.pCalendar
    Public Shared WithEvents pFlag As MUSTER.BusinessLogic.pFlag
    Dim LocalUserSettings As Microsoft.Win32.Registry
    'Private Const strMusterShare As String = "MUSTER"
    Public Shared DomainUser As New System.Security.Principal.WindowsPrincipal(WindowsIdentity.GetCurrent())

    Public Shared WithEvents pOwn As MUSTER.BusinessLogic.pOwner
    Public Shared WithEvents pLetter As MUSTER.BusinessLogic.pLetter

    Public Shared pEntity As MUSTER.BusinessLogic.pEntity
    Private strOwnerQSKeyWord As String
    Public Shared pConStruct As MUSTER.BusinessLogic.pContactStruct
    Friend WithEvents cComSearch As CompanySearchResults
    Public Shared WordApp As Word.Application
    Public WithEvents pInspec As MUSTER.BusinessLogic.pInspection
    Dim bolpnlVisibilityChanged As Boolean = False
    Dim returnVal As String = String.Empty

    Dim switchingUser As Boolean = False

    Private RestoreThread As System.Threading.Thread
    Private UnzipThread As System.Threading.Thread

    Private WithEvents oRestore As SQLDMO.RestoreClass
    Private bJustClosedAll As Boolean = False
    Private WithEvents oSQLServer As SQLDMO.SQLServer
    Private aRestoreParameters() As String
    Public WithEvents ProgressScreen As ProgressScreen
    Public WithEvents TicklerScreen As TickerScreen
    Private aunzipParameters() As String

#End Region

#Region "progress events"

    Event StartProgressScreen(ByVal title As String, ByVal maxValue As Long, ByVal value As Long, ByVal msg As String, ByVal code() As String)
    Event StartTicklerScreen(ByVal id As String)
    Event FireProgressMessage(ByVal msg As String, ByVal num As Long, ByVal code As String)
    Event FireCloseProgressScreen()


#End Region

#Region "User Defined Variables - Calendar"
    Friend dsCommon As DataSet
    Friend dsOnLoad As DataSet
    Private tableStyle As DataGridTableStyle
    Private filterString As String
    Private completedTaskFilterString As String
    Private rOption As String
    Private filterStartMonth As String
    Private filterEndMonth As String
    Private filterStartWeek As String
    Private filterEndWeek As String
    Private filterEndDay As String
    Private dateSelected As Boolean
    Private nGridHeight As Integer
    Private nGridWidth As Integer
    Dim bolCalLoad As Boolean = False
    Private bolAllowDateChangedEvent As Boolean = False
#End Region

#Region "User Defined Variables - Flags"
    Public flagBarometerEntityID, flagBarometerEntityType, flagBarometerOwnerID, flagBarometerEventID, flagBarometerEventType As Integer
    Public flagBarometerModule, flagBarometerParentFormText As String
    Private WithEvents SF As ShowFlags
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()

        MyBase.New()

        bolLoadingForm = True
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not AppUser Is Nothing Then
            AppUser.LogExit(MyGUID.ToString)
        End If
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
    Public WithEvents mnuMain As System.Windows.Forms.MainMenu
    Public HoldClosing As Boolean = False
    Public LoggedIn As Boolean = False

    Friend WithEvents x As OwnerSearchResults
    Friend WithEvents CtxMenuRightbarMode As System.Windows.Forms.ContextMenu
    Friend WithEvents pnlRightContainer As System.Windows.Forms.Panel
    Private RightPanelDockFlag As Boolean
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents tabPageCalendar As System.Windows.Forms.TabPage
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents btnQuickOwnerSearch As System.Windows.Forms.Button
    Friend WithEvents txtOwnerQSKeyword As System.Windows.Forms.TextBox
    Friend WithEvents LinkLabel6 As System.Windows.Forms.LinkLabel
    Friend WithEvents mnItemTechnical As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemCAE As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemRegistration As System.Windows.Forms.MenuItem
    Friend WithEvents btnAdvancedSearch As System.Windows.Forms.Button
    Friend WithEvents lblOwner As System.Windows.Forms.Label
    Friend WithEvents lblOwnerName As System.Windows.Forms.Label
    Friend WithEvents lblFacility As System.Windows.Forms.Label
    Friend WithEvents lblFacilityName As System.Windows.Forms.Label
    Friend WithEvents lblFacilityAddress As System.Windows.Forms.Label
    Friend WithEvents grpBxQuickSearch As System.Windows.Forms.GroupBox
    Friend WithEvents pnlCommonReferenceArea As System.Windows.Forms.Panel
    Friend WithEvents mnuItemHelp As System.Windows.Forms.MenuItem
    Friend WithEvents btnDock As System.Windows.Forms.Button
    Friend WithEvents mnuItemServices As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemOwner As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemAddOwner As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemModifyContact As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemModifyPersonalInfo As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemFacility As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemAddFacility As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemTransferOwnership As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemCAP As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemWindows As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemAboutMuster As System.Windows.Forms.MenuItem
    Friend WithEvents pnlRightMost As System.Windows.Forms.Panel
    Friend WithEvents tbCtrlRightPane As System.Windows.Forms.TabControl
    Friend WithEvents lblDueToMe As System.Windows.Forms.Label
    Friend WithEvents lblToDos As System.Windows.Forms.Label
    Friend WithEvents pnlQuickSearch As System.Windows.Forms.Panel
    Friend WithEvents cmbQuickSearchFilter As System.Windows.Forms.ComboBox
    Friend WithEvents mnItemReports As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemRegReports As System.Windows.Forms.MenuItem
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents txtUserWelcome As System.Windows.Forms.TextBox
    Friend WithEvents pnlBottomRight As System.Windows.Forms.Panel
    Friend WithEvents lblOwnerAddress As System.Windows.Forms.Label
    Friend WithEvents lblModuleComboDesc As System.Windows.Forms.Label
    Friend WithEvents lblSearchValueDesc As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCalNewEntry As System.Windows.Forms.Button
    Friend WithEvents pnlFilter As System.Windows.Forms.Panel
    Friend WithEvents grpBxCalFilter As System.Windows.Forms.GroupBox
    Friend WithEvents lblViewentries As System.Windows.Forms.Label
    Friend WithEvents chkToDoCompItems As System.Windows.Forms.CheckBox
    Friend WithEvents rdCalDay As System.Windows.Forms.RadioButton
    Friend WithEvents rdCalWeek As System.Windows.Forms.RadioButton
    Friend WithEvents rdCalMonth As System.Windows.Forms.RadioButton
    Friend WithEvents pnlCalendar As System.Windows.Forms.Panel
    Friend WithEvents calTechnicalMonth As System.Windows.Forms.MonthCalendar
    Friend WithEvents pnlToDo As System.Windows.Forms.Panel
    Friend WithEvents pnlCalToDoGrid As System.Windows.Forms.Panel
    Friend WithEvents chkDueToMeCompItems As System.Windows.Forms.CheckBox
    Friend WithEvents pnlCalToDoButtons As System.Windows.Forms.Panel
    Friend WithEvents btnToDoMarkCompleted As System.Windows.Forms.Button
    Friend WithEvents btnToDoDelete As System.Windows.Forms.Button
    Friend WithEvents pnlDueToMeCaption As System.Windows.Forms.Panel
    Friend WithEvents pnlDueToMeGrid As System.Windows.Forms.Panel
    Friend WithEvents btnDueToMeDelete As System.Windows.Forms.Button
    Friend WithEvents btnDueToMeMarkComp As System.Windows.Forms.Button
    Friend WithEvents cmbViewCalEntries As System.Windows.Forms.ComboBox
    Friend WithEvents pnlDueToMeButtons As System.Windows.Forms.Panel
    Friend WithEvents dgCalToDo As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents dgDueToMe As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblOwnerInfo As System.Windows.Forms.Label
    Friend WithEvents lblFacilityInfo As System.Windows.Forms.Label
    Friend WithEvents lnkLblNextForm As System.Windows.Forms.LinkLabel
    Friend WithEvents lnklblPrevForm As System.Windows.Forms.LinkLabel
    Friend WithEvents btnOwnerFlag As System.Windows.Forms.Button
    Friend WithEvents btnFacFlag As System.Windows.Forms.Button
    Friend WithEvents btnFeeFlag As System.Windows.Forms.Button
    Friend WithEvents btnCandEFlag As System.Windows.Forms.Button
    Friend WithEvents btnFinFlag As System.Windows.Forms.Button
    Friend WithEvents btnLUSFlag As System.Windows.Forms.Button
    Friend WithEvents btnInsFlag As System.Windows.Forms.Button
    Friend WithEvents btnCloFlag As System.Windows.Forms.Button
    Friend WithEvents btnCalToDoModify As System.Windows.Forms.Button
    Friend WithEvents btnCalDueToMeModify As System.Windows.Forms.Button
    Friend WithEvents mnuLetters As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCAPMonthly As System.Windows.Forms.MenuItem
    Friend WithEvents pnlFormClose As System.Windows.Forms.Panel
    Friend WithEvents mnuItemOwnerDocsPhotos As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemFacilityDocsPhotos As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCAPYearly As System.Windows.Forms.MenuItem
    Friend WithEvents pnlSearchCollapse As System.Windows.Forms.Panel
    Friend WithEvents lblSearchCollapse As System.Windows.Forms.Label
    Friend WithEvents lblCollapseText As System.Windows.Forms.Label
    Friend WithEvents btnFirmLicFlag As System.Windows.Forms.Button
    Friend WithEvents btnIndvLicFlag As System.Windows.Forms.Button
    Friend WithEvents mnuItemPrevFacs As System.Windows.Forms.MenuItem
    Friend WithEvents lblModuleName As System.Windows.Forms.Label
    Friend WithEvents ctxMenuHeaderColor As System.Windows.Forms.ContextMenu
    Friend WithEvents mnItemModuleColor As System.Windows.Forms.MenuItem
    Friend WithEvents ColorPicker As System.Windows.Forms.ColorDialog
    Friend WithEvents MenuItemPreOwners As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemClosure As System.Windows.Forms.MenuItem
    Public WithEvents mnuCancelAllOverDue As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItmContact As System.Windows.Forms.MenuItem
    Friend WithEvents mnuContactReconciliation As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemFinancial As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemRemSysHist As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemCompanyModule As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemCompany As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemAddCompany As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemLicensees As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemAddLicensee As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemAddComplianceManager As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemProviders As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemCourses As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemExitApp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemInspector As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemFees As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSubItemAdminServices As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSubItemFees As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemLicenseeMgmt As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemCAEManagement As System.Windows.Forms.MenuItem
    Friend WithEvents btnCloseForm As System.Windows.Forms.Button
    Friend WithEvents lblFacilityID As System.Windows.Forms.Label
    Friend WithEvents mnuFinRollover As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFinRolloverPO As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemModifyFacility As System.Windows.Forms.MenuItem
    Public WithEvents cmbSearchModule As System.Windows.Forms.ComboBox

    Friend WithEvents lblVersion As System.Windows.Forms.Label
    Friend WithEvents mnuItemSyncManualDocs As System.Windows.Forms.MenuItem
    Friend WithEvents calendarRefreshTimer As System.Windows.Forms.Timer
    Friend WithEvents mnuCAPPreMonthly As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemComplianceLetter As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemNonComplianceLetter As System.Windows.Forms.MenuItem
    Friend WithEvents mnuTOSILetter As System.Windows.Forms.MenuItem
    Friend WithEvents mnItemCloseAll As System.Windows.Forms.MenuItem
    Friend WithEvents btnCloseAll As System.Windows.Forms.Button
    Friend WithEvents mnItemSyncDB As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemOpenFacPicsFolder As System.Windows.Forms.MenuItem
    Friend WithEvents btnTickler As System.Windows.Forms.Button
    Friend WithEvents MnHistoryItems As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSpecialCAPMonthly As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCAPCurrent As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MusterContainer))
        Me.mnuMain = New System.Windows.Forms.MainMenu
        Me.mnuItemServices = New System.Windows.Forms.MenuItem
        Me.mnuItemRegistration = New System.Windows.Forms.MenuItem
        Me.mnuItemOwner = New System.Windows.Forms.MenuItem
        Me.mnuItemAddOwner = New System.Windows.Forms.MenuItem
        Me.mnuItemModifyContact = New System.Windows.Forms.MenuItem
        Me.mnItemModifyPersonalInfo = New System.Windows.Forms.MenuItem
        Me.mnuItemOwnerDocsPhotos = New System.Windows.Forms.MenuItem
        Me.mnuItemPrevFacs = New System.Windows.Forms.MenuItem
        Me.mnItemFacility = New System.Windows.Forms.MenuItem
        Me.mnItemAddFacility = New System.Windows.Forms.MenuItem
        Me.mnItemModifyFacility = New System.Windows.Forms.MenuItem
        Me.mnuItemFacilityDocsPhotos = New System.Windows.Forms.MenuItem
        Me.MenuItemPreOwners = New System.Windows.Forms.MenuItem
        Me.mnItemTransferOwnership = New System.Windows.Forms.MenuItem
        Me.mnuItemComplianceLetter = New System.Windows.Forms.MenuItem
        Me.mnuItemNonComplianceLetter = New System.Windows.Forms.MenuItem
        Me.mnuTOSILetter = New System.Windows.Forms.MenuItem

        Me.mnItemCAP = New System.Windows.Forms.MenuItem
        Me.mnuCAPPreMonthly = New System.Windows.Forms.MenuItem
        Me.mnuSpecialCAPMonthly = New System.Windows.Forms.MenuItem
        Me.mnuCAPMonthly = New System.Windows.Forms.MenuItem
        Me.mnuCAPYearly = New System.Windows.Forms.MenuItem
        Me.mnuCAPCurrent = New System.Windows.Forms.MenuItem
        Me.mnuItemFees = New System.Windows.Forms.MenuItem
        Me.mnuSubItemAdminServices = New System.Windows.Forms.MenuItem
        Me.mnuSubItemFees = New System.Windows.Forms.MenuItem
        Me.mnuItemCompanyModule = New System.Windows.Forms.MenuItem
        Me.mnuItemAddCompany = New System.Windows.Forms.MenuItem
        Me.mnuItemAddLicensee = New System.Windows.Forms.MenuItem
        Me.mnuItemAddComplianceManager = New System.Windows.Forms.MenuItem
        Me.mnuItemLicenseeMgmt = New System.Windows.Forms.MenuItem
        Me.mnuItemCompany = New System.Windows.Forms.MenuItem
        Me.mnuItemLicensees = New System.Windows.Forms.MenuItem
        Me.mnuItemProviders = New System.Windows.Forms.MenuItem
        Me.mnuItemCourses = New System.Windows.Forms.MenuItem
        Me.mnItemTechnical = New System.Windows.Forms.MenuItem
        Me.mnuItemRemSysHist = New System.Windows.Forms.MenuItem
        Me.mnItemCAE = New System.Windows.Forms.MenuItem
        Me.mnuItemCAEManagement = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuItemInspector = New System.Windows.Forms.MenuItem
        Me.mnuItemClosure = New System.Windows.Forms.MenuItem
        Me.mnuCancelAllOverDue = New System.Windows.Forms.MenuItem
        Me.mnuItmContact = New System.Windows.Forms.MenuItem
        Me.mnuContactReconciliation = New System.Windows.Forms.MenuItem
        Me.mnuItemFinancial = New System.Windows.Forms.MenuItem
        Me.mnuFinRollover = New System.Windows.Forms.MenuItem
        Me.mnuFinRolloverPO = New System.Windows.Forms.MenuItem
        Me.mnuItemExitApp = New System.Windows.Forms.MenuItem
        Me.mnItemReports = New System.Windows.Forms.MenuItem
        Me.mnItemRegReports = New System.Windows.Forms.MenuItem
        Me.mnuLetters = New System.Windows.Forms.MenuItem
        Me.mnuItemSyncManualDocs = New System.Windows.Forms.MenuItem
        Me.mnuItemOpenFacPicsFolder = New System.Windows.Forms.MenuItem
        Me.mnItemWindows = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MnHistoryItems = New System.Windows.Forms.MenuItem
        Me.mnuItemHelp = New System.Windows.Forms.MenuItem
        Me.mnItemAboutMuster = New System.Windows.Forms.MenuItem
        Me.mnItemCloseAll = New System.Windows.Forms.MenuItem
        Me.mnItemSyncDB = New System.Windows.Forms.MenuItem
        Me.pnlRightMost = New System.Windows.Forms.Panel
        Me.btnDock = New System.Windows.Forms.Button
        Me.CtxMenuRightbarMode = New System.Windows.Forms.ContextMenu
        Me.pnlRightContainer = New System.Windows.Forms.Panel
        Me.tbCtrlRightPane = New System.Windows.Forms.TabControl
        Me.tabPageCalendar = New System.Windows.Forms.TabPage
        Me.pnlFormClose = New System.Windows.Forms.Panel
        Me.pnlDueToMeButtons = New System.Windows.Forms.Panel
        Me.btnCalDueToMeModify = New System.Windows.Forms.Button
        Me.btnDueToMeDelete = New System.Windows.Forms.Button
        Me.btnDueToMeMarkComp = New System.Windows.Forms.Button
        Me.pnlDueToMeGrid = New System.Windows.Forms.Panel
        Me.dgDueToMe = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlDueToMeCaption = New System.Windows.Forms.Panel
        Me.lblDueToMe = New System.Windows.Forms.Label
        Me.chkDueToMeCompItems = New System.Windows.Forms.CheckBox
        Me.pnlCalToDoButtons = New System.Windows.Forms.Panel
        Me.btnCalToDoModify = New System.Windows.Forms.Button
        Me.btnToDoDelete = New System.Windows.Forms.Button
        Me.btnToDoMarkCompleted = New System.Windows.Forms.Button
        Me.pnlCalToDoGrid = New System.Windows.Forms.Panel
        Me.dgCalToDo = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlToDo = New System.Windows.Forms.Panel
        Me.lblToDos = New System.Windows.Forms.Label
        Me.chkToDoCompItems = New System.Windows.Forms.CheckBox
        Me.pnlFilter = New System.Windows.Forms.Panel
        Me.grpBxCalFilter = New System.Windows.Forms.GroupBox
        Me.rdCalMonth = New System.Windows.Forms.RadioButton
        Me.rdCalWeek = New System.Windows.Forms.RadioButton
        Me.rdCalDay = New System.Windows.Forms.RadioButton
        Me.cmbViewCalEntries = New System.Windows.Forms.ComboBox
        Me.btnCalNewEntry = New System.Windows.Forms.Button
        Me.lblViewentries = New System.Windows.Forms.Label
        Me.pnlCalendar = New System.Windows.Forms.Panel
        Me.calTechnicalMonth = New System.Windows.Forms.MonthCalendar
        Me.btnTickler = New System.Windows.Forms.Button
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.pnlQuickSearch = New System.Windows.Forms.Panel
        Me.grpBxQuickSearch = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblSearchValueDesc = New System.Windows.Forms.Label
        Me.lblModuleComboDesc = New System.Windows.Forms.Label
        Me.cmbSearchModule = New System.Windows.Forms.ComboBox
        Me.btnAdvancedSearch = New System.Windows.Forms.Button
        Me.cmbQuickSearchFilter = New System.Windows.Forms.ComboBox
        Me.btnQuickOwnerSearch = New System.Windows.Forms.Button
        Me.txtOwnerQSKeyword = New System.Windows.Forms.TextBox
        Me.pnlCommonReferenceArea = New System.Windows.Forms.Panel
        Me.ctxMenuHeaderColor = New System.Windows.Forms.ContextMenu
        Me.mnItemModuleColor = New System.Windows.Forms.MenuItem
        Me.lblOwnerInfo = New System.Windows.Forms.Label
        Me.lblFacilityID = New System.Windows.Forms.Label
        Me.lblModuleName = New System.Windows.Forms.Label
        Me.lnkLblNextForm = New System.Windows.Forms.LinkLabel
        Me.lnklblPrevForm = New System.Windows.Forms.LinkLabel
        Me.lblFacilityInfo = New System.Windows.Forms.Label
        Me.lblFacilityAddress = New System.Windows.Forms.Label
        Me.lblOwnerAddress = New System.Windows.Forms.Label
        Me.lblOwner = New System.Windows.Forms.Label
        Me.lblOwnerName = New System.Windows.Forms.Label
        Me.lblFacility = New System.Windows.Forms.Label
        Me.LinkLabel6 = New System.Windows.Forms.LinkLabel
        Me.lblFacilityName = New System.Windows.Forms.Label
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.lblVersion = New System.Windows.Forms.Label
        Me.btnCloseForm = New System.Windows.Forms.Button
        Me.pnlBottomRight = New System.Windows.Forms.Panel
        Me.btnIndvLicFlag = New System.Windows.Forms.Button
        Me.btnInsFlag = New System.Windows.Forms.Button
        Me.btnCandEFlag = New System.Windows.Forms.Button
        Me.btnFinFlag = New System.Windows.Forms.Button
        Me.btnLUSFlag = New System.Windows.Forms.Button
        Me.btnFirmLicFlag = New System.Windows.Forms.Button
        Me.btnCloFlag = New System.Windows.Forms.Button
        Me.btnFeeFlag = New System.Windows.Forms.Button
        Me.btnFacFlag = New System.Windows.Forms.Button
        Me.btnOwnerFlag = New System.Windows.Forms.Button
        Me.txtUserWelcome = New System.Windows.Forms.TextBox
        Me.btnCloseAll = New System.Windows.Forms.Button
        Me.pnlSearchCollapse = New System.Windows.Forms.Panel
        Me.lblCollapseText = New System.Windows.Forms.Label
        Me.lblSearchCollapse = New System.Windows.Forms.Label
        Me.ColorPicker = New System.Windows.Forms.ColorDialog
        Me.calendarRefreshTimer = New System.Windows.Forms.Timer(Me.components)
        Me.pnlRightMost.SuspendLayout()
        Me.pnlRightContainer.SuspendLayout()
        Me.tbCtrlRightPane.SuspendLayout()
        Me.tabPageCalendar.SuspendLayout()
        Me.pnlDueToMeButtons.SuspendLayout()
        Me.pnlDueToMeGrid.SuspendLayout()
        CType(Me.dgDueToMe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlDueToMeCaption.SuspendLayout()
        Me.pnlCalToDoButtons.SuspendLayout()
        Me.pnlCalToDoGrid.SuspendLayout()
        CType(Me.dgCalToDo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlToDo.SuspendLayout()
        Me.pnlFilter.SuspendLayout()
        Me.grpBxCalFilter.SuspendLayout()
        Me.pnlCalendar.SuspendLayout()
        Me.pnlQuickSearch.SuspendLayout()
        Me.grpBxQuickSearch.SuspendLayout()
        Me.pnlCommonReferenceArea.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        Me.pnlBottomRight.SuspendLayout()
        Me.pnlSearchCollapse.SuspendLayout()
        Me.SuspendLayout()
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemServices, Me.mnItemReports, Me.mnItemWindows, Me.MnHistoryItems, Me.mnuItemHelp})
        '
        'mnuItemServices
        '
        Me.mnuItemServices.Index = 0
        Me.mnuItemServices.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemRegistration, Me.mnItemCAP, Me.mnuItemFees, Me.mnuItemCompanyModule, Me.mnItemTechnical, Me.mnItemCAE, Me.mnuItemInspector, Me.mnuItemClosure, Me.mnuItmContact, Me.mnuItemFinancial, Me.mnuItemExitApp})
        Me.mnuItemServices.Text = "&Services"
        '
        'mnuItemRegistration
        '
        Me.mnuItemRegistration.Index = 0
        Me.mnuItemRegistration.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemOwner, Me.mnItemFacility, Me.mnItemTransferOwnership, Me.mnuItemComplianceLetter, Me.mnuItemNonComplianceLetter, Me.mnuTOSILetter})
        Me.mnuItemRegistration.Text = "&Registration"
        '
        'mnuItemOwner
        '
        Me.mnuItemOwner.Index = 0
        Me.mnuItemOwner.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemAddOwner, Me.mnuItemModifyContact, Me.mnItemModifyPersonalInfo, Me.mnuItemOwnerDocsPhotos, Me.mnuItemPrevFacs})
        Me.mnuItemOwner.Text = "&Owner"
        '
        'mnuItemAddOwner
        '
        Me.mnuItemAddOwner.Index = 0
        Me.mnuItemAddOwner.Text = "&Add Owner"
        '
        'mnuItemModifyContact
        '
        Me.mnuItemModifyContact.Enabled = False
        Me.mnuItemModifyContact.Index = 1
        Me.mnuItemModifyContact.Text = "&Modify Contact Information"
        '
        'mnItemModifyPersonalInfo
        '
        Me.mnItemModifyPersonalInfo.Enabled = False
        Me.mnItemModifyPersonalInfo.Index = 2
        Me.mnItemModifyPersonalInfo.Text = "&Modify Personal Information"
        '
        'mnuItemOwnerDocsPhotos
        '
        Me.mnuItemOwnerDocsPhotos.Index = 3
        Me.mnuItemOwnerDocsPhotos.Text = "&Images"
        '
        'mnuItemPrevFacs
        '
        Me.mnuItemPrevFacs.Index = 4
        Me.mnuItemPrevFacs.Text = "&Previous Facilities"
        '
        'mnItemFacility
        '
        Me.mnItemFacility.Index = 1
        Me.mnItemFacility.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnItemAddFacility, Me.mnItemModifyFacility, Me.mnuItemFacilityDocsPhotos, Me.MenuItemPreOwners})
        Me.mnItemFacility.Text = "&Facility"
        '
        'mnItemAddFacility
        '
        Me.mnItemAddFacility.Index = 0
        Me.mnItemAddFacility.Text = "&Add New Facility"
        '
        'mnItemModifyFacility
        '
        Me.mnItemModifyFacility.Enabled = False
        Me.mnItemModifyFacility.Index = 1
        Me.mnItemModifyFacility.Text = "&Modify Facility"
        '
        'mnuItemFacilityDocsPhotos
        '
        Me.mnuItemFacilityDocsPhotos.Index = 2
        Me.mnuItemFacilityDocsPhotos.Text = "&Images"
        '
        'MenuItemPreOwners
        '
        Me.MenuItemPreOwners.Index = 3
        Me.MenuItemPreOwners.Text = "&Previous Owners"
        '
        'mnItemTransferOwnership
        '
        Me.mnItemTransferOwnership.Index = 2
        Me.mnItemTransferOwnership.Text = "&Transfer Ownership"
        '
        'mnuItemComplianceLetter
        '
        Me.mnuItemComplianceLetter.Index = 3
        Me.mnuItemComplianceLetter.Text = "&Compliance Letter"
        '
        'mnuItemNonComplianceLetter
        '
        Me.mnuItemNonComplianceLetter.Index = 4
        Me.mnuItemNonComplianceLetter.Text = "&Non Compliance Letter"
        Me.mnuItemNonComplianceLetter.Visible = False
        '
        'mnuTOSILetter
        '
        Me.mnuTOSILetter.Index = 5
        Me.mnuTOSILetter.Text = "&TOS-I Letter"

        '
        'mnItemCAP
        '
        Me.mnItemCAP.Index = 1
        Me.mnItemCAP.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCAPPreMonthly, Me.mnuSpecialCAPMonthly, Me.mnuCAPMonthly, Me.mnuCAPYearly, Me.mnuCAPCurrent})
        Me.mnItemCAP.Text = "&CAP Program"
        '
        'mnuCAPPreMonthly
        '
        Me.mnuCAPPreMonthly.Index = 0
        Me.mnuCAPPreMonthly.Text = "&Pre-Monthly CAP Processing"
        '
        'mnuSpecialCAPMonthly
        '
        Me.mnuSpecialCAPMonthly.Index = 1
        Me.mnuSpecialCAPMonthly.Text = "&Special Monthly Cap Processing"
        '
        'mnuCAPMonthly
        '
        Me.mnuCAPMonthly.Index = 2
        Me.mnuCAPMonthly.Text = "&Monthly CAP Processing"
        '
        'mnuCAPYearly
        '
        Me.mnuCAPYearly.Index = 3
        Me.mnuCAPYearly.Text = "&Yearly CAP Processing"
        '
        'mnuCAPCurrent
        '
        Me.mnuCAPCurrent.Index = 4
        Me.mnuCAPCurrent.Text = "&Current CAP Processing"
        '
        'mnuItemFees
        '
        Me.mnuItemFees.Index = 2
        Me.mnuItemFees.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSubItemAdminServices, Me.mnuSubItemFees})
        Me.mnuItemFees.Text = "&Fees"
        '
        'mnuSubItemAdminServices
        '
        Me.mnuSubItemAdminServices.Index = 0
        Me.mnuSubItemAdminServices.Text = "&Adminstrative Services"
        '
        'mnuSubItemFees
        '
        Me.mnuSubItemFees.Index = 1
        Me.mnuSubItemFees.Text = "&Fees"
        '
        'mnuItemCompanyModule
        '
        Me.mnuItemCompanyModule.Index = 3
        Me.mnuItemCompanyModule.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemAddCompany, Me.mnuItemAddLicensee, Me.mnuItemLicenseeMgmt, Me.mnuItemCompany, Me.mnuItemLicensees, Me.mnuItemProviders, Me.mnuItemCourses, Me.mnuItemAddComplianceManager})
        Me.mnuItemCompanyModule.Text = "C&ompany"
        '
        'mnuItemAddCompany
        '
        Me.mnuItemAddCompany.Index = 0
        Me.mnuItemAddCompany.Text = "Add New &Company"
        '
        'mnuItemAddLicensee
        '
        Me.mnuItemAddLicensee.Index = 1
        Me.mnuItemAddLicensee.Text = "Add New &Licensee"
        '
        'mnuItemLicenseeMgmt
        '
        Me.mnuItemLicenseeMgmt.Index = 2
        Me.mnuItemLicenseeMgmt.Text = "Licensee &Management"
        '
        'mnuItemCompany
        '
        Me.mnuItemCompany.Index = 3
        Me.mnuItemCompany.Text = "&Company"
        Me.mnuItemCompany.Visible = False
        '
        'mnuItemLicensees
        '
        Me.mnuItemLicensees.Index = 4
        Me.mnuItemLicensees.Text = "&Licensees"
        Me.mnuItemLicensees.Visible = False
        '
        'mnuItemProviders
        '
        Me.mnuItemProviders.Index = 5
        Me.mnuItemProviders.Text = "Manage &Providers"
        Me.mnuItemProviders.Visible = False
        '
        'mnuItemCourses
        '
        Me.mnuItemCourses.Index = 6
        Me.mnuItemCourses.Text = "Manage &Courses"
        Me.mnuItemCourses.Visible = False
        '
        'mnuItemAddComplianceManager
        '
        Me.mnuItemAddComplianceManager.Index = 7
        Me.mnuItemAddComplianceManager.Text = "Add New Compliance &Manager"
        '
        'mnItemTechnical
        '
        Me.mnItemTechnical.Index = 4
        Me.mnItemTechnical.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemRemSysHist})
        Me.mnItemTechnical.Text = "&Technical"
        '
        'mnuItemRemSysHist
        '
        Me.mnuItemRemSysHist.Index = 0
        Me.mnuItemRemSysHist.Text = "&Rem Sys History"
        '
        'mnItemCAE
        '
        Me.mnItemCAE.Index = 5
        Me.mnItemCAE.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemCAEManagement, Me.MenuItem3})
        Me.mnItemCAE.Text = "C && &E"
        '
        'mnuItemCAEManagement
        '
        Me.mnuItemCAEManagement.Index = 0
        Me.mnuItemCAEManagement.Text = "&Management"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "&Assignments"
        '
        'mnuItemInspector
        '
        Me.mnuItemInspector.Index = 6
        Me.mnuItemInspector.Text = "&Inspector"
        '
        'mnuItemClosure
        '
        Me.mnuItemClosure.Index = 7
        Me.mnuItemClosure.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCancelAllOverDue})
        Me.mnuItemClosure.Text = "&Closure"
        '
        'mnuCancelAllOverDue
        '
        Me.mnuCancelAllOverDue.Enabled = False
        Me.mnuCancelAllOverDue.Index = 0
        Me.mnuCancelAllOverDue.Text = "&Cancel All Overdue"
        '
        'mnuItmContact
        '
        Me.mnuItmContact.Enabled = False
        Me.mnuItmContact.Index = 8
        Me.mnuItmContact.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuContactReconciliation})
        Me.mnuItmContact.Text = "C&ontact"
        '
        'mnuContactReconciliation
        '
        Me.mnuContactReconciliation.Index = 0
        Me.mnuContactReconciliation.Text = "&Reconciliation"
        '
        'mnuItemFinancial
        '
        Me.mnuItemFinancial.Index = 9
        Me.mnuItemFinancial.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFinRollover, Me.mnuFinRolloverPO})
        Me.mnuItemFinancial.Text = "&Financial"
        '
        'mnuFinRollover
        '
        Me.mnuFinRollover.Index = 0
        Me.mnuFinRollover.Text = "&Rollover && ZeroOut"
        '
        'mnuFinRolloverPO
        '
        Me.mnuFinRolloverPO.Index = 1
        Me.mnuFinRolloverPO.Text = "&New Rollover PO Numbers"
        '
        'mnuItemExitApp
        '
        Me.mnuItemExitApp.Index = 10
        Me.mnuItemExitApp.Text = "E&xit Application"
        '
        'mnItemReports
        '
        Me.mnItemReports.Index = 1
        Me.mnItemReports.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnItemRegReports, Me.mnuLetters, Me.mnuItemSyncManualDocs, Me.mnuItemOpenFacPicsFolder})
        Me.mnItemReports.Text = "&Utilities"
        '
        'mnItemRegReports
        '
        Me.mnItemRegReports.Index = 0
        Me.mnItemRegReports.Text = "&Reports"
        '
        'mnuLetters
        '
        Me.mnuLetters.Index = 1
        Me.mnuLetters.Text = "&Documents List"
        '
        'mnuItemSyncManualDocs
        '
        Me.mnuItemSyncManualDocs.Index = 2
        Me.mnuItemSyncManualDocs.Text = "Sync Manual Docs"
        '
        'mnuItemOpenFacPicsFolder
        '
        Me.mnuItemOpenFacPicsFolder.Index = 3
        Me.mnuItemOpenFacPicsFolder.Text = "&Open FacPics Folder"
        '
        'mnItemWindows
        '
        Me.mnItemWindows.Index = 2
        Me.mnItemWindows.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem1})
        Me.mnItemWindows.ShowShortcut = False
        Me.mnItemWindows.Text = "&Windows"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "-"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.Text = "&Close All"
        '
        'MnHistoryItems
        '
        Me.MnHistoryItems.Index = 3
        Me.MnHistoryItems.Text = "H&istory"
        Me.MnHistoryItems.Visible = False
        '
        'mnuItemHelp
        '
        Me.mnuItemHelp.Index = 4
        Me.mnuItemHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnItemAboutMuster, Me.mnItemCloseAll, Me.mnItemSyncDB})
        Me.mnuItemHelp.Text = "&Help"
        '
        'mnItemAboutMuster
        '
        Me.mnItemAboutMuster.Index = 0
        Me.mnItemAboutMuster.Text = "&About MUSTER"
        '
        'mnItemCloseAll
        '
        Me.mnItemCloseAll.Enabled = False
        Me.mnItemCloseAll.Index = 1
        Me.mnItemCloseAll.Text = "&Close All Windows"
        '
        'mnItemSyncDB
        '
        Me.mnItemSyncDB.Index = 2
        Me.mnItemSyncDB.Text = "&Sync Database"
        '
        'pnlRightMost
        '
        Me.pnlRightMost.BackColor = System.Drawing.SystemColors.Info
        Me.pnlRightMost.Controls.Add(Me.btnDock)
        Me.pnlRightMost.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnlRightMost.DockPadding.Right = 16
        Me.pnlRightMost.Location = New System.Drawing.Point(1014, 0)
        Me.pnlRightMost.Name = "pnlRightMost"
        Me.pnlRightMost.Size = New System.Drawing.Size(14, 625)
        Me.pnlRightMost.TabIndex = 14
        '
        'btnDock
        '
        Me.btnDock.ContextMenu = Me.CtxMenuRightbarMode
        Me.btnDock.Location = New System.Drawing.Point(0, 0)
        Me.btnDock.Name = "btnDock"
        Me.btnDock.Size = New System.Drawing.Size(14, 24)
        Me.btnDock.TabIndex = 16
        Me.btnDock.Text = ">"
        '
        'pnlRightContainer
        '
        Me.pnlRightContainer.AutoScroll = True
        Me.pnlRightContainer.Controls.Add(Me.tbCtrlRightPane)
        Me.pnlRightContainer.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnlRightContainer.Location = New System.Drawing.Point(718, 0)
        Me.pnlRightContainer.Name = "pnlRightContainer"
        Me.pnlRightContainer.Size = New System.Drawing.Size(296, 625)
        Me.pnlRightContainer.TabIndex = 17
        '
        'tbCtrlRightPane
        '
        Me.tbCtrlRightPane.Controls.Add(Me.tabPageCalendar)
        Me.tbCtrlRightPane.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlRightPane.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlRightPane.Name = "tbCtrlRightPane"
        Me.tbCtrlRightPane.SelectedIndex = 0
        Me.tbCtrlRightPane.Size = New System.Drawing.Size(296, 625)
        Me.tbCtrlRightPane.TabIndex = 1
        '
        'tabPageCalendar
        '
        Me.tabPageCalendar.Controls.Add(Me.pnlFormClose)
        Me.tabPageCalendar.Controls.Add(Me.pnlDueToMeButtons)
        Me.tabPageCalendar.Controls.Add(Me.pnlDueToMeGrid)
        Me.tabPageCalendar.Controls.Add(Me.pnlDueToMeCaption)
        Me.tabPageCalendar.Controls.Add(Me.pnlCalToDoButtons)
        Me.tabPageCalendar.Controls.Add(Me.pnlCalToDoGrid)
        Me.tabPageCalendar.Controls.Add(Me.pnlToDo)
        Me.tabPageCalendar.Controls.Add(Me.pnlFilter)
        Me.tabPageCalendar.Controls.Add(Me.pnlCalendar)
        Me.tabPageCalendar.Location = New System.Drawing.Point(4, 23)
        Me.tabPageCalendar.Name = "tabPageCalendar"
        Me.tabPageCalendar.Size = New System.Drawing.Size(288, 598)
        Me.tabPageCalendar.TabIndex = 0
        Me.tabPageCalendar.Text = "Calendar"
        Me.tabPageCalendar.ToolTipText = "Calendar View"
        '
        'pnlFormClose
        '
        Me.pnlFormClose.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFormClose.Location = New System.Drawing.Point(0, 584)
        Me.pnlFormClose.Name = "pnlFormClose"
        Me.pnlFormClose.Size = New System.Drawing.Size(288, 32)
        Me.pnlFormClose.TabIndex = 28
        '
        'pnlDueToMeButtons
        '
        Me.pnlDueToMeButtons.Controls.Add(Me.btnCalDueToMeModify)
        Me.pnlDueToMeButtons.Controls.Add(Me.btnDueToMeDelete)
        Me.pnlDueToMeButtons.Controls.Add(Me.btnDueToMeMarkComp)
        Me.pnlDueToMeButtons.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlDueToMeButtons.Location = New System.Drawing.Point(0, 561)
        Me.pnlDueToMeButtons.Name = "pnlDueToMeButtons"
        Me.pnlDueToMeButtons.Size = New System.Drawing.Size(288, 23)
        Me.pnlDueToMeButtons.TabIndex = 27
        '
        'btnCalDueToMeModify
        '
        Me.btnCalDueToMeModify.Location = New System.Drawing.Point(111, 0)
        Me.btnCalDueToMeModify.Name = "btnCalDueToMeModify"
        Me.btnCalDueToMeModify.TabIndex = 2
        Me.btnCalDueToMeModify.Text = "Modify"
        '
        'btnDueToMeDelete
        '
        Me.btnDueToMeDelete.Location = New System.Drawing.Point(193, 1)
        Me.btnDueToMeDelete.Name = "btnDueToMeDelete"
        Me.btnDueToMeDelete.TabIndex = 1
        Me.btnDueToMeDelete.Text = "Delete"
        '
        'btnDueToMeMarkComp
        '
        Me.btnDueToMeMarkComp.Location = New System.Drawing.Point(10, 1)
        Me.btnDueToMeMarkComp.Name = "btnDueToMeMarkComp"
        Me.btnDueToMeMarkComp.Size = New System.Drawing.Size(96, 23)
        Me.btnDueToMeMarkComp.TabIndex = 0
        Me.btnDueToMeMarkComp.Text = "Mark Completed"
        '
        'pnlDueToMeGrid
        '
        Me.pnlDueToMeGrid.Controls.Add(Me.dgDueToMe)
        Me.pnlDueToMeGrid.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlDueToMeGrid.Location = New System.Drawing.Point(0, 448)
        Me.pnlDueToMeGrid.Name = "pnlDueToMeGrid"
        Me.pnlDueToMeGrid.Size = New System.Drawing.Size(288, 113)
        Me.pnlDueToMeGrid.TabIndex = 26
        '
        'dgDueToMe
        '
        Me.dgDueToMe.Cursor = System.Windows.Forms.Cursors.Default
        Me.dgDueToMe.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.dgDueToMe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgDueToMe.Location = New System.Drawing.Point(0, 0)
        Me.dgDueToMe.Name = "dgDueToMe"
        Me.dgDueToMe.Size = New System.Drawing.Size(288, 113)
        Me.dgDueToMe.TabIndex = 0
        '
        'pnlDueToMeCaption
        '
        Me.pnlDueToMeCaption.Controls.Add(Me.lblDueToMe)
        Me.pnlDueToMeCaption.Controls.Add(Me.chkDueToMeCompItems)
        Me.pnlDueToMeCaption.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlDueToMeCaption.Location = New System.Drawing.Point(0, 423)
        Me.pnlDueToMeCaption.Name = "pnlDueToMeCaption"
        Me.pnlDueToMeCaption.Size = New System.Drawing.Size(288, 25)
        Me.pnlDueToMeCaption.TabIndex = 22
        '
        'lblDueToMe
        '
        Me.lblDueToMe.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDueToMe.Location = New System.Drawing.Point(2, 0)
        Me.lblDueToMe.Name = "lblDueToMe"
        Me.lblDueToMe.Size = New System.Drawing.Size(112, 17)
        Me.lblDueToMe.TabIndex = 39
        Me.lblDueToMe.Text = "Items Due to Me"
        Me.lblDueToMe.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkDueToMeCompItems
        '
        Me.chkDueToMeCompItems.Location = New System.Drawing.Point(148, 0)
        Me.chkDueToMeCompItems.Name = "chkDueToMeCompItems"
        Me.chkDueToMeCompItems.Size = New System.Drawing.Size(140, 24)
        Me.chkDueToMeCompItems.TabIndex = 42
        Me.chkDueToMeCompItems.Text = "Show Completed Items"
        '
        'pnlCalToDoButtons
        '
        Me.pnlCalToDoButtons.Controls.Add(Me.btnCalToDoModify)
        Me.pnlCalToDoButtons.Controls.Add(Me.btnToDoDelete)
        Me.pnlCalToDoButtons.Controls.Add(Me.btnToDoMarkCompleted)
        Me.pnlCalToDoButtons.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCalToDoButtons.Location = New System.Drawing.Point(0, 400)
        Me.pnlCalToDoButtons.Name = "pnlCalToDoButtons"
        Me.pnlCalToDoButtons.Size = New System.Drawing.Size(288, 23)
        Me.pnlCalToDoButtons.TabIndex = 24
        '
        'btnCalToDoModify
        '
        Me.btnCalToDoModify.Location = New System.Drawing.Point(111, 0)
        Me.btnCalToDoModify.Name = "btnCalToDoModify"
        Me.btnCalToDoModify.TabIndex = 2
        Me.btnCalToDoModify.Text = "Modify"
        '
        'btnToDoDelete
        '
        Me.btnToDoDelete.Location = New System.Drawing.Point(194, 0)
        Me.btnToDoDelete.Name = "btnToDoDelete"
        Me.btnToDoDelete.TabIndex = 1
        Me.btnToDoDelete.Text = "Delete"
        '
        'btnToDoMarkCompleted
        '
        Me.btnToDoMarkCompleted.Location = New System.Drawing.Point(9, 0)
        Me.btnToDoMarkCompleted.Name = "btnToDoMarkCompleted"
        Me.btnToDoMarkCompleted.Size = New System.Drawing.Size(96, 23)
        Me.btnToDoMarkCompleted.TabIndex = 0
        Me.btnToDoMarkCompleted.Text = "Mark Completed"
        '
        'pnlCalToDoGrid
        '
        Me.pnlCalToDoGrid.Controls.Add(Me.dgCalToDo)
        Me.pnlCalToDoGrid.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCalToDoGrid.Location = New System.Drawing.Point(0, 271)
        Me.pnlCalToDoGrid.Name = "pnlCalToDoGrid"
        Me.pnlCalToDoGrid.Size = New System.Drawing.Size(288, 129)
        Me.pnlCalToDoGrid.TabIndex = 23
        '
        'dgCalToDo
        '
        Me.dgCalToDo.Cursor = System.Windows.Forms.Cursors.Default
        Me.dgCalToDo.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.dgCalToDo.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.dgCalToDo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgCalToDo.Location = New System.Drawing.Point(0, 0)
        Me.dgCalToDo.Name = "dgCalToDo"
        Me.dgCalToDo.Size = New System.Drawing.Size(288, 129)
        Me.dgCalToDo.TabIndex = 0
        '
        'pnlToDo
        '
        Me.pnlToDo.Controls.Add(Me.lblToDos)
        Me.pnlToDo.Controls.Add(Me.chkToDoCompItems)
        Me.pnlToDo.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlToDo.Location = New System.Drawing.Point(0, 248)
        Me.pnlToDo.Name = "pnlToDo"
        Me.pnlToDo.Size = New System.Drawing.Size(288, 23)
        Me.pnlToDo.TabIndex = 22
        '
        'lblToDos
        '
        Me.lblToDos.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblToDos.Location = New System.Drawing.Point(2, 0)
        Me.lblToDos.Name = "lblToDos"
        Me.lblToDos.Size = New System.Drawing.Size(92, 17)
        Me.lblToDos.TabIndex = 37
        Me.lblToDos.Text = "List of To Do"
        Me.lblToDos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkToDoCompItems
        '
        Me.chkToDoCompItems.Location = New System.Drawing.Point(148, 0)
        Me.chkToDoCompItems.Name = "chkToDoCompItems"
        Me.chkToDoCompItems.Size = New System.Drawing.Size(140, 24)
        Me.chkToDoCompItems.TabIndex = 41
        Me.chkToDoCompItems.Text = "Show Completed Items"
        '
        'pnlFilter
        '
        Me.pnlFilter.Controls.Add(Me.grpBxCalFilter)
        Me.pnlFilter.Controls.Add(Me.cmbViewCalEntries)
        Me.pnlFilter.Controls.Add(Me.btnCalNewEntry)
        Me.pnlFilter.Controls.Add(Me.lblViewentries)
        Me.pnlFilter.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFilter.Location = New System.Drawing.Point(0, 168)
        Me.pnlFilter.Name = "pnlFilter"
        Me.pnlFilter.Size = New System.Drawing.Size(288, 80)
        Me.pnlFilter.TabIndex = 20
        '
        'grpBxCalFilter
        '
        Me.grpBxCalFilter.Controls.Add(Me.rdCalMonth)
        Me.grpBxCalFilter.Controls.Add(Me.rdCalWeek)
        Me.grpBxCalFilter.Controls.Add(Me.rdCalDay)
        Me.grpBxCalFilter.Location = New System.Drawing.Point(5, 32)
        Me.grpBxCalFilter.Name = "grpBxCalFilter"
        Me.grpBxCalFilter.Size = New System.Drawing.Size(184, 43)
        Me.grpBxCalFilter.TabIndex = 39
        Me.grpBxCalFilter.TabStop = False
        Me.grpBxCalFilter.Text = "Filter"
        '
        'rdCalMonth
        '
        Me.rdCalMonth.Location = New System.Drawing.Point(128, 14)
        Me.rdCalMonth.Name = "rdCalMonth"
        Me.rdCalMonth.Size = New System.Drawing.Size(54, 20)
        Me.rdCalMonth.TabIndex = 2
        Me.rdCalMonth.Text = "Month"
        '
        'rdCalWeek
        '
        Me.rdCalWeek.Location = New System.Drawing.Point(59, 14)
        Me.rdCalWeek.Name = "rdCalWeek"
        Me.rdCalWeek.Size = New System.Drawing.Size(59, 20)
        Me.rdCalWeek.TabIndex = 1
        Me.rdCalWeek.Text = "Week"
        '
        'rdCalDay
        '
        Me.rdCalDay.Location = New System.Drawing.Point(8, 14)
        Me.rdCalDay.Name = "rdCalDay"
        Me.rdCalDay.Size = New System.Drawing.Size(48, 20)
        Me.rdCalDay.TabIndex = 0
        Me.rdCalDay.Text = "Day"
        '
        'cmbViewCalEntries
        '
        Me.cmbViewCalEntries.Enabled = False
        Me.cmbViewCalEntries.Items.AddRange(New Object() {"Sandy"})
        Me.cmbViewCalEntries.Location = New System.Drawing.Point(190, 51)
        Me.cmbViewCalEntries.Name = "cmbViewCalEntries"
        Me.cmbViewCalEntries.Size = New System.Drawing.Size(96, 22)
        Me.cmbViewCalEntries.TabIndex = 38
        '
        'btnCalNewEntry
        '
        Me.btnCalNewEntry.Location = New System.Drawing.Point(56, 8)
        Me.btnCalNewEntry.Name = "btnCalNewEntry"
        Me.btnCalNewEntry.Size = New System.Drawing.Size(64, 23)
        Me.btnCalNewEntry.TabIndex = 37
        Me.btnCalNewEntry.Text = "New Entry"
        '
        'lblViewentries
        '
        Me.lblViewentries.Location = New System.Drawing.Point(192, 25)
        Me.lblViewentries.Name = "lblViewentries"
        Me.lblViewentries.Size = New System.Drawing.Size(90, 21)
        Me.lblViewentries.TabIndex = 40
        Me.lblViewentries.Text = "View Entries For:"
        '
        'pnlCalendar
        '
        Me.pnlCalendar.Controls.Add(Me.calTechnicalMonth)
        Me.pnlCalendar.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCalendar.Location = New System.Drawing.Point(0, 0)
        Me.pnlCalendar.Name = "pnlCalendar"
        Me.pnlCalendar.Size = New System.Drawing.Size(288, 168)
        Me.pnlCalendar.TabIndex = 21
        '
        'calTechnicalMonth
        '
        Me.calTechnicalMonth.Dock = System.Windows.Forms.DockStyle.Top
        Me.calTechnicalMonth.Location = New System.Drawing.Point(0, 0)
        Me.calTechnicalMonth.Name = "calTechnicalMonth"
        Me.calTechnicalMonth.TabIndex = 36
        '
        'btnTickler
        '
        Me.btnTickler.Location = New System.Drawing.Point(320, 1)
        Me.btnTickler.Name = "btnTickler"
        Me.btnTickler.Size = New System.Drawing.Size(296, 23)
        Me.btnTickler.TabIndex = 29
        Me.btnTickler.Text = "There are no unread tickler messages."
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Date Due"
        Me.ColumnHeader3.Width = 82
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Document"
        Me.ColumnHeader4.Width = 110
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Date Received"
        Me.ColumnHeader1.Width = 85
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Document"
        Me.ColumnHeader2.Width = 107
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Splitter1.Location = New System.Drawing.Point(710, 0)
        Me.Splitter1.MinExtra = 304
        Me.Splitter1.MinSize = 304
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(8, 625)
        Me.Splitter1.TabIndex = 22
        Me.Splitter1.TabStop = False
        '
        'pnlQuickSearch
        '
        Me.pnlQuickSearch.Controls.Add(Me.grpBxQuickSearch)
        Me.pnlQuickSearch.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlQuickSearch.Location = New System.Drawing.Point(0, 16)
        Me.pnlQuickSearch.Name = "pnlQuickSearch"
        Me.pnlQuickSearch.Size = New System.Drawing.Size(710, 56)
        Me.pnlQuickSearch.TabIndex = 25
        '
        'grpBxQuickSearch
        '
        Me.grpBxQuickSearch.Controls.Add(Me.Label1)
        Me.grpBxQuickSearch.Controls.Add(Me.lblSearchValueDesc)
        Me.grpBxQuickSearch.Controls.Add(Me.lblModuleComboDesc)
        Me.grpBxQuickSearch.Controls.Add(Me.cmbSearchModule)
        Me.grpBxQuickSearch.Controls.Add(Me.btnAdvancedSearch)
        Me.grpBxQuickSearch.Controls.Add(Me.cmbQuickSearchFilter)
        Me.grpBxQuickSearch.Controls.Add(Me.btnQuickOwnerSearch)
        Me.grpBxQuickSearch.Controls.Add(Me.txtOwnerQSKeyword)
        Me.grpBxQuickSearch.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpBxQuickSearch.Location = New System.Drawing.Point(0, 0)
        Me.grpBxQuickSearch.Name = "grpBxQuickSearch"
        Me.grpBxQuickSearch.Size = New System.Drawing.Size(710, 75)
        Me.grpBxQuickSearch.TabIndex = 1
        Me.grpBxQuickSearch.TabStop = False
        Me.grpBxQuickSearch.Text = "Quick Search"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(312, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Search By"
        '
        'lblSearchValueDesc
        '
        Me.lblSearchValueDesc.Location = New System.Drawing.Point(56, 14)
        Me.lblSearchValueDesc.Name = "lblSearchValueDesc"
        Me.lblSearchValueDesc.Size = New System.Drawing.Size(56, 16)
        Me.lblSearchValueDesc.TabIndex = 6
        Me.lblSearchValueDesc.Text = "Look For"
        '
        'lblModuleComboDesc
        '
        Me.lblModuleComboDesc.Location = New System.Drawing.Point(186, 14)
        Me.lblModuleComboDesc.Name = "lblModuleComboDesc"
        Me.lblModuleComboDesc.Size = New System.Drawing.Size(56, 16)
        Me.lblModuleComboDesc.TabIndex = 5
        Me.lblModuleComboDesc.Text = "Module"
        '
        'cmbSearchModule
        '
        Me.cmbSearchModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSearchModule.Location = New System.Drawing.Point(152, 31)
        Me.cmbSearchModule.MaxDropDownItems = 9
        Me.cmbSearchModule.Name = "cmbSearchModule"
        Me.cmbSearchModule.Size = New System.Drawing.Size(120, 22)
        Me.cmbSearchModule.TabIndex = 1
        '
        'btnAdvancedSearch
        '
        Me.btnAdvancedSearch.Location = New System.Drawing.Point(466, 31)
        Me.btnAdvancedSearch.Name = "btnAdvancedSearch"
        Me.btnAdvancedSearch.Size = New System.Drawing.Size(108, 24)
        Me.btnAdvancedSearch.TabIndex = 4
        Me.btnAdvancedSearch.Text = "Advanced Search"
        '
        'cmbQuickSearchFilter
        '
        Me.cmbQuickSearchFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbQuickSearchFilter.Location = New System.Drawing.Point(280, 31)
        Me.cmbQuickSearchFilter.Name = "cmbQuickSearchFilter"
        Me.cmbQuickSearchFilter.Size = New System.Drawing.Size(120, 22)
        Me.cmbQuickSearchFilter.TabIndex = 2
        '
        'btnQuickOwnerSearch
        '
        Me.btnQuickOwnerSearch.Location = New System.Drawing.Point(405, 31)
        Me.btnQuickOwnerSearch.Name = "btnQuickOwnerSearch"
        Me.btnQuickOwnerSearch.Size = New System.Drawing.Size(56, 24)
        Me.btnQuickOwnerSearch.TabIndex = 3
        Me.btnQuickOwnerSearch.Text = "Search"
        '
        'txtOwnerQSKeyword
        '
        Me.txtOwnerQSKeyword.Location = New System.Drawing.Point(8, 31)
        Me.txtOwnerQSKeyword.Name = "txtOwnerQSKeyword"
        Me.txtOwnerQSKeyword.Size = New System.Drawing.Size(144, 20)
        Me.txtOwnerQSKeyword.TabIndex = 0
        Me.txtOwnerQSKeyword.Text = ""
        Me.txtOwnerQSKeyword.WordWrap = False
        '
        'pnlCommonReferenceArea
        '
        Me.pnlCommonReferenceArea.BackColor = System.Drawing.Color.Transparent
        Me.pnlCommonReferenceArea.ContextMenu = Me.ctxMenuHeaderColor
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblOwnerInfo)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblFacilityID)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblModuleName)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lnkLblNextForm)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lnklblPrevForm)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblFacilityInfo)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblFacilityAddress)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblOwnerAddress)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblOwner)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblOwnerName)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblFacility)
        Me.pnlCommonReferenceArea.Controls.Add(Me.LinkLabel6)
        Me.pnlCommonReferenceArea.Controls.Add(Me.lblFacilityName)
        Me.pnlCommonReferenceArea.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCommonReferenceArea.Location = New System.Drawing.Point(0, 72)
        Me.pnlCommonReferenceArea.Name = "pnlCommonReferenceArea"
        Me.pnlCommonReferenceArea.Size = New System.Drawing.Size(710, 40)
        Me.pnlCommonReferenceArea.TabIndex = 24
        Me.pnlCommonReferenceArea.Visible = False
        '
        'ctxMenuHeaderColor
        '
        Me.ctxMenuHeaderColor.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnItemModuleColor})
        '
        'mnItemModuleColor
        '
        Me.mnItemModuleColor.Index = 0
        Me.mnItemModuleColor.Text = "&Module Color"
        '
        'lblOwnerInfo
        '
        Me.lblOwnerInfo.Location = New System.Drawing.Point(224, 6)
        Me.lblOwnerInfo.Name = "lblOwnerInfo"
        Me.lblOwnerInfo.Size = New System.Drawing.Size(224, 32)
        Me.lblOwnerInfo.TabIndex = 10
        Me.lblOwnerInfo.Visible = False
        '
        'lblFacilityID
        '
        Me.lblFacilityID.BackColor = System.Drawing.Color.Transparent
        Me.lblFacilityID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityID.Location = New System.Drawing.Point(528, 6)
        Me.lblFacilityID.Name = "lblFacilityID"
        Me.lblFacilityID.Size = New System.Drawing.Size(72, 18)
        Me.lblFacilityID.TabIndex = 1036
        Me.lblFacilityID.Visible = False
        '
        'lblModuleName
        '
        Me.lblModuleName.BackColor = System.Drawing.Color.Transparent
        Me.lblModuleName.Font = New System.Drawing.Font("Arial", 14.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModuleName.Location = New System.Drawing.Point(8, 2)
        Me.lblModuleName.Name = "lblModuleName"
        Me.lblModuleName.Size = New System.Drawing.Size(120, 24)
        Me.lblModuleName.TabIndex = 1035
        Me.lblModuleName.Text = "Registration"
        Me.lblModuleName.Visible = False
        '
        'lnkLblNextForm
        '
        Me.lnkLblNextForm.AutoSize = True
        Me.lnkLblNextForm.BackColor = System.Drawing.Color.Transparent
        Me.lnkLblNextForm.Location = New System.Drawing.Point(488, 22)
        Me.lnkLblNextForm.Name = "lnkLblNextForm"
        Me.lnkLblNextForm.Size = New System.Drawing.Size(24, 16)
        Me.lnkLblNextForm.TabIndex = 1034
        Me.lnkLblNextForm.TabStop = True
        Me.lnkLblNextForm.Text = "Add"
        Me.lnkLblNextForm.Visible = False
        '
        'lnklblPrevForm
        '
        Me.lnklblPrevForm.AutoSize = True
        Me.lnklblPrevForm.BackColor = System.Drawing.Color.Transparent
        Me.lnklblPrevForm.Location = New System.Drawing.Point(166, 21)
        Me.lnklblPrevForm.Name = "lnklblPrevForm"
        Me.lnklblPrevForm.Size = New System.Drawing.Size(24, 16)
        Me.lnklblPrevForm.TabIndex = 1033
        Me.lnklblPrevForm.TabStop = True
        Me.lnklblPrevForm.Text = "Add"
        Me.lnklblPrevForm.Visible = False
        '
        'lblFacilityInfo
        '
        Me.lblFacilityInfo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityInfo.Location = New System.Drawing.Point(528, 22)
        Me.lblFacilityInfo.Name = "lblFacilityInfo"
        Me.lblFacilityInfo.Size = New System.Drawing.Size(248, 16)
        Me.lblFacilityInfo.TabIndex = 11
        Me.lblFacilityInfo.Visible = False
        '
        'lblFacilityAddress
        '
        Me.lblFacilityAddress.AutoSize = True
        Me.lblFacilityAddress.Location = New System.Drawing.Point(320, 24)
        Me.lblFacilityAddress.Name = "lblFacilityAddress"
        Me.lblFacilityAddress.Size = New System.Drawing.Size(0, 16)
        Me.lblFacilityAddress.TabIndex = 9
        Me.lblFacilityAddress.Visible = False
        '
        'lblOwnerAddress
        '
        Me.lblOwnerAddress.AutoSize = True
        Me.lblOwnerAddress.Location = New System.Drawing.Point(24, 24)
        Me.lblOwnerAddress.Name = "lblOwnerAddress"
        Me.lblOwnerAddress.Size = New System.Drawing.Size(0, 16)
        Me.lblOwnerAddress.TabIndex = 8
        Me.lblOwnerAddress.Visible = False
        '
        'lblOwner
        '
        Me.lblOwner.AutoSize = True
        Me.lblOwner.BackColor = System.Drawing.Color.Transparent
        Me.lblOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwner.Location = New System.Drawing.Point(141, 6)
        Me.lblOwner.Name = "lblOwner"
        Me.lblOwner.Size = New System.Drawing.Size(50, 18)
        Me.lblOwner.TabIndex = 2
        Me.lblOwner.Text = "Owner:"
        Me.lblOwner.Visible = False
        '
        'lblOwnerName
        '
        Me.lblOwnerName.AutoSize = True
        Me.lblOwnerName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerName.Location = New System.Drawing.Point(73, 6)
        Me.lblOwnerName.Name = "lblOwnerName"
        Me.lblOwnerName.Size = New System.Drawing.Size(0, 18)
        Me.lblOwnerName.TabIndex = 2
        Me.lblOwnerName.Visible = False
        '
        'lblFacility
        '
        Me.lblFacility.AutoSize = True
        Me.lblFacility.BackColor = System.Drawing.Color.Transparent
        Me.lblFacility.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacility.Location = New System.Drawing.Point(464, 6)
        Me.lblFacility.Name = "lblFacility"
        Me.lblFacility.Size = New System.Drawing.Size(53, 18)
        Me.lblFacility.TabIndex = 2
        Me.lblFacility.Text = "Facility:"
        Me.lblFacility.Visible = False
        '
        'LinkLabel6
        '
        Me.LinkLabel6.AutoSize = True
        Me.LinkLabel6.Location = New System.Drawing.Point(32, 24)
        Me.LinkLabel6.Name = "LinkLabel6"
        Me.LinkLabel6.Size = New System.Drawing.Size(0, 16)
        Me.LinkLabel6.TabIndex = 6
        '
        'lblFacilityName
        '
        Me.lblFacilityName.AutoSize = True
        Me.lblFacilityName.Location = New System.Drawing.Point(376, 8)
        Me.lblFacilityName.Name = "lblFacilityName"
        Me.lblFacilityName.Size = New System.Drawing.Size(0, 16)
        Me.lblFacilityName.TabIndex = 8
        Me.lblFacilityName.Visible = False
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.lblVersion)
        Me.pnlBottom.Controls.Add(Me.btnCloseForm)
        Me.pnlBottom.Controls.Add(Me.pnlBottomRight)
        Me.pnlBottom.Controls.Add(Me.txtUserWelcome)
        Me.pnlBottom.Controls.Add(Me.btnCloseAll)
        Me.pnlBottom.Controls.Add(Me.btnTickler)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 625)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(1028, 24)
        Me.pnlBottom.TabIndex = 27
        '
        'lblVersion
        '
        Me.lblVersion.BackColor = System.Drawing.Color.Transparent
        Me.lblVersion.Font = New System.Drawing.Font("Times New Roman", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVersion.Location = New System.Drawing.Point(552, 3)
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(63, 18)
        Me.lblVersion.TabIndex = 14
        Me.lblVersion.Visible = False
        '
        'btnCloseForm
        '
        Me.btnCloseForm.Enabled = False
        Me.btnCloseForm.Location = New System.Drawing.Point(160, 1)
        Me.btnCloseForm.Name = "btnCloseForm"
        Me.btnCloseForm.TabIndex = 11
        Me.btnCloseForm.Text = "Close"
        '
        'pnlBottomRight
        '
        Me.pnlBottomRight.Controls.Add(Me.btnIndvLicFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnInsFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnCandEFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnFinFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnLUSFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnFirmLicFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnCloFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnFeeFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnFacFlag)
        Me.pnlBottomRight.Controls.Add(Me.btnOwnerFlag)
        Me.pnlBottomRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnlBottomRight.Location = New System.Drawing.Point(616, 0)
        Me.pnlBottomRight.Name = "pnlBottomRight"
        Me.pnlBottomRight.Size = New System.Drawing.Size(412, 24)
        Me.pnlBottomRight.TabIndex = 10
        '
        'btnIndvLicFlag
        '
        Me.btnIndvLicFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnIndvLicFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnIndvLicFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnIndvLicFlag.Location = New System.Drawing.Point(365, 0)
        Me.btnIndvLicFlag.Name = "btnIndvLicFlag"
        Me.btnIndvLicFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnIndvLicFlag.TabIndex = 25
        Me.btnIndvLicFlag.Text = "Indv"
        Me.btnIndvLicFlag.Visible = False
        '
        'btnInsFlag
        '
        Me.btnInsFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnInsFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnInsFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnInsFlag.Location = New System.Drawing.Point(285, 0)
        Me.btnInsFlag.Name = "btnInsFlag"
        Me.btnInsFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnInsFlag.TabIndex = 24
        Me.btnInsFlag.Text = "Insp"
        Me.btnInsFlag.Visible = False
        '
        'btnCandEFlag
        '
        Me.btnCandEFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnCandEFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCandEFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCandEFlag.Location = New System.Drawing.Point(245, 0)
        Me.btnCandEFlag.Name = "btnCandEFlag"
        Me.btnCandEFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnCandEFlag.TabIndex = 23
        Me.btnCandEFlag.Text = "C&&E"
        Me.btnCandEFlag.Visible = False
        '
        'btnFinFlag
        '
        Me.btnFinFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnFinFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFinFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFinFlag.Location = New System.Drawing.Point(205, 0)
        Me.btnFinFlag.Name = "btnFinFlag"
        Me.btnFinFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnFinFlag.TabIndex = 22
        Me.btnFinFlag.Text = "Fin"
        Me.btnFinFlag.Visible = False
        '
        'btnLUSFlag
        '
        Me.btnLUSFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnLUSFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnLUSFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLUSFlag.Location = New System.Drawing.Point(165, 0)
        Me.btnLUSFlag.Name = "btnLUSFlag"
        Me.btnLUSFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnLUSFlag.TabIndex = 21
        Me.btnLUSFlag.Text = "Lust"
        Me.btnLUSFlag.Visible = False
        '
        'btnFirmLicFlag
        '
        Me.btnFirmLicFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnFirmLicFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFirmLicFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFirmLicFlag.Location = New System.Drawing.Point(325, 0)
        Me.btnFirmLicFlag.Name = "btnFirmLicFlag"
        Me.btnFirmLicFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnFirmLicFlag.TabIndex = 20
        Me.btnFirmLicFlag.Text = "Firm"
        Me.btnFirmLicFlag.Visible = False
        '
        'btnCloFlag
        '
        Me.btnCloFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnCloFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCloFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCloFlag.Location = New System.Drawing.Point(125, 0)
        Me.btnCloFlag.Name = "btnCloFlag"
        Me.btnCloFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnCloFlag.TabIndex = 19
        Me.btnCloFlag.Text = "Clos"
        Me.btnCloFlag.Visible = False
        '
        'btnFeeFlag
        '
        Me.btnFeeFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnFeeFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFeeFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFeeFlag.Location = New System.Drawing.Point(85, 0)
        Me.btnFeeFlag.Name = "btnFeeFlag"
        Me.btnFeeFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnFeeFlag.TabIndex = 18
        Me.btnFeeFlag.Text = "Fee"
        Me.btnFeeFlag.Visible = False
        '
        'btnFacFlag
        '
        Me.btnFacFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnFacFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnFacFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFacFlag.Location = New System.Drawing.Point(45, 0)
        Me.btnFacFlag.Name = "btnFacFlag"
        Me.btnFacFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnFacFlag.TabIndex = 17
        Me.btnFacFlag.Text = "Reg"
        Me.btnFacFlag.Visible = False
        '
        'btnOwnerFlag
        '
        Me.btnOwnerFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerFlag.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnOwnerFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOwnerFlag.Location = New System.Drawing.Point(5, 0)
        Me.btnOwnerFlag.Name = "btnOwnerFlag"
        Me.btnOwnerFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnOwnerFlag.TabIndex = 16
        Me.btnOwnerFlag.Text = "Own"
        Me.btnOwnerFlag.Visible = False
        '
        'txtUserWelcome
        '
        Me.txtUserWelcome.BackColor = System.Drawing.SystemColors.Control
        Me.txtUserWelcome.Location = New System.Drawing.Point(3, 2)
        Me.txtUserWelcome.Name = "txtUserWelcome"
        Me.txtUserWelcome.Size = New System.Drawing.Size(157, 20)
        Me.txtUserWelcome.TabIndex = 8
        Me.txtUserWelcome.Text = ""
        '
        'btnCloseAll
        '
        Me.btnCloseAll.Enabled = False
        Me.btnCloseAll.Location = New System.Drawing.Point(240, 1)
        Me.btnCloseAll.Name = "btnCloseAll"
        Me.btnCloseAll.TabIndex = 11
        Me.btnCloseAll.Text = "Close All"
        '
        'pnlSearchCollapse
        '
        Me.pnlSearchCollapse.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.pnlSearchCollapse.Controls.Add(Me.lblCollapseText)
        Me.pnlSearchCollapse.Controls.Add(Me.lblSearchCollapse)
        Me.pnlSearchCollapse.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlSearchCollapse.Location = New System.Drawing.Point(0, 0)
        Me.pnlSearchCollapse.Name = "pnlSearchCollapse"
        Me.pnlSearchCollapse.Size = New System.Drawing.Size(710, 16)
        Me.pnlSearchCollapse.TabIndex = 24
        '
        'lblCollapseText
        '
        Me.lblCollapseText.BackColor = System.Drawing.Color.Transparent
        Me.lblCollapseText.Font = New System.Drawing.Font("Arial", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCollapseText.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCollapseText.Location = New System.Drawing.Point(24, 0)
        Me.lblCollapseText.Name = "lblCollapseText"
        Me.lblCollapseText.Size = New System.Drawing.Size(312, 18)
        Me.lblCollapseText.TabIndex = 1
        Me.lblCollapseText.Text = "Collapse Search"
        '
        'lblSearchCollapse
        '
        Me.lblSearchCollapse.BackColor = System.Drawing.Color.WhiteSmoke
        Me.lblSearchCollapse.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchCollapse.Location = New System.Drawing.Point(0, 0)
        Me.lblSearchCollapse.Name = "lblSearchCollapse"
        Me.lblSearchCollapse.Size = New System.Drawing.Size(16, 18)
        Me.lblSearchCollapse.TabIndex = 0
        Me.lblSearchCollapse.Text = "-"
        '
        'calendarRefreshTimer
        '
        Me.calendarRefreshTimer.Interval = 1800000
        '
        'MusterContainer
        '
        Me.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1028, 649)
        Me.Controls.Add(Me.pnlCommonReferenceArea)
        Me.Controls.Add(Me.pnlQuickSearch)
        Me.Controls.Add(Me.pnlSearchCollapse)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.pnlRightContainer)
        Me.Controls.Add(Me.pnlRightMost)
        Me.Controls.Add(Me.pnlBottom)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Menu = Me.mnuMain
        Me.Name = "MusterContainer"
        Me.Text = "MUSTER"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlRightMost.ResumeLayout(False)
        Me.pnlRightContainer.ResumeLayout(False)
        Me.tbCtrlRightPane.ResumeLayout(False)
        Me.tabPageCalendar.ResumeLayout(False)
        Me.pnlDueToMeButtons.ResumeLayout(False)
        Me.pnlDueToMeGrid.ResumeLayout(False)
        CType(Me.dgDueToMe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlDueToMeCaption.ResumeLayout(False)
        Me.pnlCalToDoButtons.ResumeLayout(False)
        Me.pnlCalToDoGrid.ResumeLayout(False)
        CType(Me.dgCalToDo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlToDo.ResumeLayout(False)
        Me.pnlFilter.ResumeLayout(False)
        Me.grpBxCalFilter.ResumeLayout(False)
        Me.pnlCalendar.ResumeLayout(False)
        Me.pnlQuickSearch.ResumeLayout(False)
        Me.grpBxQuickSearch.ResumeLayout(False)
        Me.pnlCommonReferenceArea.ResumeLayout(False)
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlBottomRight.ResumeLayout(False)
        Me.pnlSearchCollapse.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Advanced Search"
    Private Sub btnAdvancedSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdvancedSearch.Click
        If IsNothing(frmAdvanceSearch) Then
            Try
                frmAdvanceSearch = New AdvancedSearch
                AddHandler frmAdvanceSearch.Closing, AddressOf frmAdvClosing
                AddHandler frmAdvanceSearch.Closed, AddressOf frmAdvclosed
            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
                Exit Sub
            End Try
            MyGUID = frmAdvanceSearch.MyGUID
            MusterContainer.AppSemaphores.Retrieve(MyGUID.ToString, "WindowName", "Advanced Search")
            'ElseIf frmAdvanceSearch.IsDisposed Then
        Else
            frmAdvanceSearch.Close()
            frmAdvanceSearch = New AdvancedSearch
            AddHandler frmAdvanceSearch.Closing, AddressOf frmAdvClosing
            AddHandler frmAdvanceSearch.Closed, AddressOf frmAdvclosed
        End If
        AddHandler frmAdvanceSearch.SearchResultSelection, AddressOf x_SearchResultSelection
        AddHandler frmAdvanceSearch.CompanySearchSelection, AddressOf Company_SearchResultSelection
        frmAdvanceSearch.WindowState = FormWindowState.Maximized
        frmAdvanceSearch.BringToFront()
        frmAdvanceSearch.MdiParent = Me
        frmAdvanceSearch.Show()
        frmAdvanceSearch.cmbFavoriteSearches.SelectedIndex = -1
    End Sub
    Private Sub frmAdvClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        frmAdvanceSearch = Nothing
    End Sub
    Private Sub frmAdvclosed(ByVal sender As Object, ByVal e As System.EventArgs)
        frmAdvanceSearch = Nothing
    End Sub
    Private Sub mnItemAdvSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If IsNothing(frmAdvanceSearch) Then
            Try
                frmAdvanceSearch = New AdvancedSearch
                AddHandler x.Closing, AddressOf frmResultClosing
                AddHandler x.Closed, AddressOf frmResultClosed
            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            End Try
        End If
        frmAdvanceSearch.WindowState = FormWindowState.Maximized
        frmAdvanceSearch.BringToFront()
        frmAdvanceSearch.MdiParent = Me
        frmAdvanceSearch.Show()
    End Sub
    Private Sub frmResultClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        x = Nothing
    End Sub
    Private Sub frmResultClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        x = Nothing
    End Sub
#End Region

#Region "Search"

    Sub InvokeMnulettersClick()
        mnuLetters.PerformClick()
    End Sub

    Sub InvokeReportClick()
        mnItemRegReports.PerformClick()
    End Sub

    Sub InvokeQuickOwneerButtonClick()
        btnQuickOwnerSearch.PerformClick()
    End Sub


    Public Sub OpenEntityFromTickler(ByVal moduleTxt As String, ByVal entityID As String, ByVal keyword As String) Handles TicklerScreen.OpenObject

        Dim actualEntityID As String = String.Empty

        If keyword.IndexOf("ReportView") <> -1 Then



            GotoModule = moduleTxt
            GotoReport = entityID

            BeginInvoke(Me.CallReportManager)


        ElseIf keyword.IndexOf("DocumentManager") <> -1 Then

            GoToUser = entityID.Substring(0, entityID.IndexOf(";"))
            GotoYear = entityID.Substring(entityID.IndexOf(";") + 1)

            BeginInvoke(Me.CallDocumentManager)

        Else
            If entityID.Length > 0 AndAlso entityID.IndexOf("--") > -1 Then

                GoToGrid = entityID.Substring(entityID.LastIndexOf("--") + 2)
                entityID = entityID.Substring(0, entityID.LastIndexOf("--"))

                actualEntityID = entityID.Substring(0, entityID.IndexOf("--"))
                GoToEntityCode = entityID.Substring(entityID.IndexOf("--") + 2)
            ElseIf entityID.Length > 0 AndAlso entityID.IndexOf(";") > -1 Then

                GoToTabPage = entityID.Substring(entityID.LastIndexOf(";") + 2)
                actualEntityID = entityID.Substring(0, entityID.LastIndexOf(";"))

            Else
                GoToGrid = String.Empty
                actualEntityID = entityID
            End If

            txtOwnerQSKeyword.Text = actualEntityID
            cmbSearchModule.Text = moduleTxt
            cmbQuickSearchFilter.Text = keyword

            BeginInvoke(CallOpenForm)

        End If




    End Sub



    Friend Sub btnQuickOwnerSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuickOwnerSearch.Click
        Try


            Me.Update()
            LockWindowUpdate(CLng(Me.Handle.ToInt64))
            bolErrorOccurred = False
            oQS = New MUSTER.BusinessLogic.pSearch
            AddHandler oQS.SearchErr, AddressOf SearchErr
            oQS.Keyword = txtOwnerQSKeyword.Text
            oQS.Module = cmbSearchModule.Text
            oQS.Filter = cmbQuickSearchFilter.Text

            Me.Cursor = Cursors.WaitCursor
            If cmbSearchModule.Text = "Company" And (cmbQuickSearchFilter.Text = "Company Name" Or cmbQuickSearchFilter.Text = "Licensee Name" Or cmbQuickSearchFilter.Text = "Manager Name") Then
                If IsNothing(cComSearch) Then
                    cComSearch = New CompanySearchResults(BootStrap._container)
                    bolErrorOccurred = False
                End If
                cComSearch.NewShowResults(oQS)
            Else
                If IsNothing(x) Then
                    x = New OwnerSearchResults(Me)
                    bolErrorOccurred = False
                    RemoveHandler oQS.SearchErr, AddressOf SearchErr
                    AddHandler x.OwnerSearchResultErr, AddressOf SearchErr
                End If
                x.NewShowResults(oQS)
                If bolErrorOccurred Then
                    txtOwnerQSKeyword.Focus()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
            LockWindowUpdate(CLng(0))
        End Try
    End Sub
    Private Sub cmbSearchModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSearchModule.SelectedIndexChanged


        If bolLoadingForm Then Exit Sub
        Dim dtTable As DataTable
        Try
            dtTable = oQS.PopulateQuickSearchFilter(cmbSearchModule.SelectedValue)
            cmbQuickSearchFilter.DataSource = dtTable
            If Not dtTable Is Nothing Then
                cmbQuickSearchFilter.ValueMember = "PROPERTY_NAME"
                cmbQuickSearchFilter.DisplayMember = "PROPERTY_NAME"
                cmbQuickSearchFilter.SelectedIndex = 1
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub RefreshQuickSearchResults()
        Dim frmChild As Form
        Dim sender As System.Object
        Dim e As System.EventArgs
        Try

            For Each frmChild In Me.MdiChildren
                If frmChild.GetType.Name = "OwnerSearchResults" Then
                    oQS.Keyword = strOwnerQSKeyWord
                    oQS.Module = cmbSearchModule.Text
                    oQS.Filter = cmbQuickSearchFilter.Text
                    Me.Cursor = Cursors.WaitCursor
                    x.NewShowResults(oQS)
                End If
            Next

        Catch ex As Exception
            Throw
        Finally
            Me.Cursor = Cursors.Default
            strOwnerQSKeyWord = String.Empty
        End Try
    End Sub
    Private Sub SearchErr(ByVal MsgStr As String, ByVal strColumnName As String, ByVal strSrc As String)
        MessageBox.Show(MsgStr)
        Select Case UCase(strColumnName)
            Case "KEYWORD"
                txtOwnerQSKeyword.Focus()
            Case "MODULE"
                cmbSearchModule.Focus()
            Case "FILTER"
                cmbQuickSearchFilter.Focus()
            Case "RESULTS"
                txtOwnerQSKeyword.Focus()
        End Select
        bolErrorOccurred = True
    End Sub
#End Region

#Region "Admin Menu"
    Private Function AddAdmin()
        'switchingUser
        For Each mnuItem As MenuItem In mnuMain.MenuItems
            If mnuItem.Text.ToUpper = "&ADMIN" Then
                Exit Function
            End If
        Next
        Dim mnuItemAdmin As New MenuItem("&Admin")
        Dim mnuItemAdminPaths As New MenuItem("Manage &File Paths", AddressOf ShowFilePathAdmin)
        Dim mnuItemAdminPaths2 As New MenuItem("&Profile Datum", AddressOf ShowProfileDatum)
        Dim mnuItemAdminPaths3 As New MenuItem("Manage &Groups", AddressOf ShowGroups)
        Dim mnuItemAdminPaths4 As New MenuItem("Manage &Users", AddressOf ShowUser)
        Dim mnuItemAdminPaths16 As New MenuItem("Manage &Module Entity Access", AddressOf ShowModuleEntityRel)
        Dim mnuItemAdminPaths5 As New MenuItem("Manage &Code Table Properties", AddressOf ShowCodeTableManager)
        Dim mnuItemAdminPaths6 As New MenuItem("Manage &Reports", AddressOf ShowReport)
        Dim mnuItemAdminPaths15 As New MenuItem("&Manage Contact Types", AddressOf ShowContactTypes)
        Dim mnuItemAdminPaths30 As New MenuItem("Switch User", AddressOf SwitchUser)
        mnuMain.MenuItems.Add(mnuItemAdmin)
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths)
#If DEBUG Then
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths2)
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths30)
#End If
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths3)
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths4)
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths16)
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths5)
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths6)
        mnuItemAdmin.MenuItems.Add(mnuItemAdminPaths15)
    End Function

    Private Sub ShowContactTypes(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim objContactTypes As New ManageContactTypes

        objContactTypes.MdiParent = Me
        objContactTypes.Show()
    End Sub

    Private Sub ShowInvoiceRequestAdmin(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim invReq As New InvoiceRequestAdmin

        invReq.ShowDialog()
        invReq.Dispose()

    End Sub

    Private Sub ShowFilePathAdmin(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oAdmin As New FilePathAdmin(Me)
        oAdmin.Show()
    End Sub
    Private Sub ShowCodeTableManager(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oPropMgr As New CodeTableManager(Me)
        oPropMgr.WindowState = FormWindowState.Maximized
        'oPropMgr.BringToFront()
        oPropMgr.Show()
    End Sub
    Private Sub ShowProfileDatum(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oProDatum As New ManageProfileData(Me)
        oProDatum.WindowState = FormWindowState.Maximized
        oProDatum.Show()

    End Sub
    Private Sub ShowGroups(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oGroups As New GroupAdmin(Me)
        oGroups.WindowState = FormWindowState.Maximized
        oGroups.Show()

    End Sub
    Private Sub ShowUser(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oUser As New UserAdmin(Me)
        oUser.WindowState = FormWindowState.Maximized
        oUser.Show()

    End Sub
    Private Sub ShowModuleEntityRel(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim moduleEntityRel As New ModuleEntityRights(Me)
        moduleEntityRel.WindowState = FormWindowState.Maximized
        moduleEntityRel.Show()
    End Sub
    Private Sub SwitchUser(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            For Each frm As Form In Me.MdiChildren
                frm.Close()
            Next

            AppUser = Nothing

            Dim ThisForm As Logon
            ThisForm = New Logon(Me)
            ThisForm.ShowDialog()
            '
            ' Check if the user actually suceeded in logging in.
            '   If the appuser object doesn't exist, then user abandonded
            '
            If AppUser Is Nothing Then
                Me.Dispose()
                Exit Sub
            End If
            If AppUser.ID = String.Empty Then
                Me.Dispose()
            End If
            bolLoadingForm = True
            switchingUser = True
            RegistrationServices_Load(sender, e)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            bolLoadingForm = False
            switchingUser = False
        End Try
    End Sub
    Private Sub ShowReport(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim rpt As New ManageReports(Me)
        rpt.WindowState = FormWindowState.Maximized
        rpt.Show()

    End Sub
#End Region

#Region "Tech Admin Menu"
    Private Sub AddTechAdmin()
        For Each mnuItem As MenuItem In mnuMain.MenuItems
            If mnuItem.Text.ToUpper = "&TECHNICAL ADMIN" Then
                Exit Sub
            End If
        Next
        Dim mnuItemTechAdmin As New MenuItem("&Technical Admin")
        Dim mnuItemAdminPaths8 As New MenuItem("&Tech Document Admin", AddressOf ShowTechDocAdmin)
        Dim mnuItemAdminPaths9 As New MenuItem("Tech &Activity Admin", AddressOf ShowTechActAdmin)
        Dim mnuItemAdminPaths15 As New MenuItem("&Manage Contact Types", AddressOf ShowContactTypes)
        mnuMain.MenuItems.Add(mnuItemTechAdmin)
        mnuItemTechAdmin.MenuItems.Add(mnuItemAdminPaths8)
        mnuItemTechAdmin.MenuItems.Add(mnuItemAdminPaths9)
        mnuItemTechAdmin.MenuItems.Add(mnuItemAdminPaths15)
    End Sub

    Private Sub ShowTechDocAdmin(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oDocList As New DocumentList
        oDocList.MdiParent = Me
        oDocList.Show()
    End Sub
    Private Sub ShowTechActAdmin(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oActList As New ActivityList

        If (AppUser.HEAD_FINANCIAL OrElse AppUser.HEAD_ADMIN) AndAlso DirectCast(sender, MenuItem).Text.IndexOf("Cost") > -1 Then
            oActList.IsFinancial = True
        ElseIf DirectCast(sender, MenuItem).Text.IndexOf("Cost") > -1 Then
            oActList.Dispose()
            MsgBox("You do not have the head financial rights to enter this form")
        End If

        oActList.MdiParent = Me
        oActList.Show()
    End Sub
#End Region

#Region "Fin Admin Menu"
    Private Sub AddFinAdmin()
        For Each mnuItem As MenuItem In mnuMain.MenuItems
            If mnuItem.Text.ToUpper = "&FINANCIAL ADMIN" Then
                Exit Sub
            End If
        Next
        Dim mnuItemFinAdmin As New MenuItem("&Financial Admin")
        Dim mnuItemAdminPaths10 As New MenuItem("Financial &Deduction Reason", AddressOf ShowFinDeductReason)
        Dim mnuItemAdminPaths11 As New MenuItem("Financial Additional &Conditions", AddressOf ShowFinAddlCondition)
        Dim mnuItemAdminPaths12 As New MenuItem("Financial &Incomplete App Reason", AddressOf ShowFinIncApp)
        Dim mnuItemAdminPaths13 As New MenuItem("Financial Condition For &Reimbursement", AddressOf ShowFinCondition)
        Dim mnuItemAdminPaths14 As New MenuItem("&Financial Activity Admin", AddressOf ShowFinActivity)
        Dim mnuItemAdminPaths16 As New MenuItem("In&voice Request Admin", AddressOf ShowInvoiceRequestAdmin)
        Dim mnuItemAdminPaths15 As New MenuItem("&Manage Contact Types", AddressOf ShowContactTypes)
        Dim mnuItemAdminPaths17 As New MenuItem("Tech &Activity Cost Admin", AddressOf ShowTechActAdmin)

        mnuMain.MenuItems.Add(mnuItemFinAdmin)
        mnuItemFinAdmin.MenuItems.Add(mnuItemAdminPaths10)
        mnuItemFinAdmin.MenuItems.Add(mnuItemAdminPaths11)
        mnuItemFinAdmin.MenuItems.Add(mnuItemAdminPaths12)
        mnuItemFinAdmin.MenuItems.Add(mnuItemAdminPaths13)
        mnuItemFinAdmin.MenuItems.Add(mnuItemAdminPaths14)
        mnuItemFinAdmin.MenuItems.Add(mnuItemAdminPaths16)
        mnuItemFinAdmin.MenuItems.Add(mnuItemAdminPaths15)
        mnuItemFinAdmin.MenuItems.Add(mnuItemAdminPaths17)

    End Sub

    Private Sub ShowFinDeductReason(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oTextAdmin As New FinTextAdmin
        oTextAdmin.MdiParent = Me
        oTextAdmin.ReasonType = 987
        oTextAdmin.Show()
    End Sub
    Private Sub ShowFinAddlCondition(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oTextAdmin As New FinTextAdmin
        oTextAdmin.MdiParent = Me
        oTextAdmin.ReasonType = 984
        oTextAdmin.Show()
    End Sub
    Private Sub ShowFinIncApp(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oTextAdmin As New FinTextAdmin
        oTextAdmin.MdiParent = Me
        oTextAdmin.ReasonType = 986
        oTextAdmin.Show()
    End Sub
    Private Sub ShowFinCondition(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oTextAdmin As New FinTextAdmin
        oTextAdmin.MdiParent = Me
        oTextAdmin.ReasonType = 983
        oTextAdmin.Show()
    End Sub
    Private Sub ShowFinActivity(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim oFinActivity As New AddModifyActivity
        oFinActivity.MdiParent = Me
        oFinActivity.Show()
    End Sub
#End Region

#Region "Company Admin"
    Private Sub AddCompanyAdmin()
        For Each mnuItem As MenuItem In mnuMain.MenuItems
            If mnuItem.Text.ToUpper = "&COMPANY ADMIN" Then
                Exit Sub
            End If
        Next
        Dim mnuItemCompanyAdmin As New MenuItem("&Company Admin")
        Dim mnuItemAdminPaths10 As New MenuItem("Manage &Providers", AddressOf mnuItemProviders_Click)
        Dim mnuItemAdminPaths11 As New MenuItem("Manage &Courses", AddressOf mnuItemCourses_Click)
        mnuMain.MenuItems.Add(mnuItemCompanyAdmin)
        mnuItemCompanyAdmin.MenuItems.Add(mnuItemAdminPaths10)
        mnuItemCompanyAdmin.MenuItems.Add(mnuItemAdminPaths11)
    End Sub

    Private Sub mnuItemProviders_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemProviders.Click
        Dim objProv As Providers
        Try
            objProv = New Providers(Me)
            objProv.Show()
            objProv.BringToFront()
            objProv.cmbProviderName.SelectedIndex = -1
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub mnuItemCourses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemCourses.Click
        Dim objCourse As Courses
        Try
            objCourse = New Courses(Me)
            objCourse.Show()
            objCourse.BringToFront()
            objCourse.cmbCourseTitle.SelectedIndex = -1
            objCourse.cmbProvider.SelectedIndex = -1
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region

#Region "Menu"
#Region "Services"
#Region "Registraton"
#Region "Owner"
    Private Sub mnuItemAddOwner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemAddOwner.Click

        Dim objRegister As New Registration
        Try
            objRegister.MdiParent = Me
            objRegister.Show()
            objRegister.BringToFront()
            objRegister.Tag = "Owner-Add"
            objRegister.tbCntrlRegistration.SelectedTab = objRegister.tbPageOwnerDetail
            objRegister.SetupAddOwner()
            HideShowOwnerInfo(True)
            objRegister.btnOwnerComment.Enabled = False
            objRegister.btnOwnerFlag.Enabled = False
            objRegister.LinkLblCAPSignup.Enabled = False

            lblOwnerInfo.Text = ""
            lblFacilityID.Text = ""
            lblFacilityInfo.Text = ""
            objRegister.NewOwner()

            ClearBarometer()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mnuItemOwnerDocsPhotos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuItemOwnerDocsPhotos.Click
        Dim DocsImages As New DocsPhotos(Me)
        'Dim dsSet As DataSet
        Dim noOfFiles As Integer
        Dim noOfImages As Integer
        Dim strEntityIds As String
        Dim nEntityId As Integer
        Dim nEntityType As Integer
        Dim strListOfEntityIds() As String
        Dim strFileEntityIds As String = ""
        'Dim dConsumer As New DocConsumer
        Dim strFilterString As String
        Dim strActiveFormValue As String
        Dim strFiles As String = ""
        Dim strFormName As String = String.Empty
        Dim TechForm As MUSTER.Technical
        Dim ClosrForm As MUSTER.Closure
        Dim FinancialForm As MUSTER.Financial
        Dim FeesForm As MUSTER.Fees
        Dim CandEForm As MUSTER.CandE
        Dim i As Integer = 0
        Dim drRow() As DataRow
        'Hard Coded
        Dim strUserId As String = AppUser.ID

        Dim ActiveGUID As Guid
        Dim ActiveOwner As Int64
        Dim RegForm As Registration
        Dim oForm As Form


        Try
            ActiveGUID = AppSemaphores.GetValuePair("0", "ActiveForm")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                Dim MyErr As ErrorReport
                MyErr = New ErrorReport(New Exception("No Documents/Images Found.You have to select a Owner", ex))
                MyErr.ShowDialog()
                Exit Sub
            End If

        End Try
        Try
            'dsSet = dConsumer.loadDocument(strUserId)
            'dsSet = pOwn.RunSQLQuery("SELECT * FROM tblSYS_DOCUMENT_MANAGER")
            ActiveOwner = MusterContainer.AppSemaphores.GetValuePair(ActiveGUID.ToString, "OwnerID", True)
            If ActiveOwner = 0 Then
                MsgBox("No Documents/Images Found.You have to select a Owner")
                Exit Sub
            Else
                '    For Each oForm In Me.MdiChildren
                '        If TypeOf (oForm) Is Registration Then
                '            RegForm = CType(oForm, Registration)
                '            RegForm.tbCntrlRegistration.SelectedTab = RegForm.tbPageOwnerDetail
                '        End If
                '        If oForm.Text.StartsWith("Registration - Owner Summary") Or oForm.Text.StartsWith("Registration - Owner Detail") Then
                '            DocsImages.Text = oForm.Text + " Documents"
                '            strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString()
                '            strEntityIds = CStr(AppSemaphores.GetValuePair(strActiveFormValue, "OwnerFacilities"))
                '            nEntityType = 6
                '        End If
                '    Next
                'End If

                'If oForm.Text.StartsWith("Registration - Owner Summary") Or oForm.Text.StartsWith("Registration - Owner Detail") Then
                Select Case cmbSearchModule.Text
                    Case "Registration"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Registration Then
                                RegForm = CType(oForm, Registration)
                                RegForm.tbCntrlRegistration.SelectedTab = RegForm.tbPageOwnerDetail
                                strFormName = RegForm.Name
                            End If
                            If oForm.Text.StartsWith("Registration - Owner Summary") Or oForm.Text.StartsWith("Registration - Owner Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString()
                                strEntityIds = CStr(AppSemaphores.GetValuePair(strActiveFormValue, "OwnerFacilities"))
                                nEntityType = 6
                            End If
                        Next
                    Case "Technical"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Technical Then
                                TechForm = CType(oForm, Technical)
                                TechForm.tbCntrlTechnical.SelectedTab = TechForm.tbPageOwnerDetail
                                strFormName = TechForm.Name
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Owner Summary") Or oForm.Text.StartsWith(strFormName + " - Owner Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString()
                                strEntityIds = CStr(AppSemaphores.GetValuePair(strActiveFormValue, "OwnerFacilities"))
                                nEntityType = 6
                            End If
                        Next
                    Case "Closure"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Closure Then
                                ClosrForm = CType(oForm, Closure)
                                ClosrForm.tbCntrlClosure.SelectedTab = ClosrForm.tbPageOwnerDetail
                                strFormName = ClosrForm.Name
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Owner Summary") Or oForm.Text.StartsWith(strFormName + " - Owner Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString()
                                strEntityIds = CStr(AppSemaphores.GetValuePair(strActiveFormValue, "OwnerFacilities"))
                                nEntityType = 6
                            End If
                        Next
                    Case "Financial"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Financial Then
                                FinancialForm = CType(oForm, Financial)
                                FinancialForm.tbCntrlFinancial.SelectedTab = FinancialForm.tbPageOwnerDetail
                                strFormName = FinancialForm.Name
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Owner Summary") Or oForm.Text.StartsWith(strFormName + " - Owner Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString()
                                strEntityIds = CStr(AppSemaphores.GetValuePair(strActiveFormValue, "OwnerFacilities"))
                                nEntityType = 6
                            End If
                        Next
                    Case "Fees"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Fees Then
                                FeesForm = CType(oForm, Fees)
                                FeesForm.tbCntrlFees.SelectedTab = FeesForm.tbPageOwnerDetail
                                strFormName = FeesForm.Name
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Owner Summary") Or oForm.Text.StartsWith(strFormName + " - Owner Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString()
                                strEntityIds = CStr(AppSemaphores.GetValuePair(strActiveFormValue, "OwnerFacilities"))
                                nEntityType = 6
                            End If
                        Next
                    Case "C & E"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is CandE Then
                                CandEForm = CType(oForm, CandE)
                                CandEForm.tbCntrlCandE.SelectedTab = CandEForm.tbPageOwnerDetail
                                strFormName = "C " + "&" + " E"
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Owner Summary") Or oForm.Text.StartsWith(strFormName + " - Owner Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString()
                                strEntityIds = CStr(AppSemaphores.GetValuePair(strActiveFormValue, "OwnerFacilities"))
                                nEntityType = 6
                            End If
                        Next
                End Select
            End If

            If oForm.Text.StartsWith(strFormName + " - Owner Summary") Or oForm.Text.StartsWith(strFormName + " - Owner Detail") Then


                If ActiveOwner <> 0 Then
                    strFilterString = " ENTITY_ID=" + ActiveOwner.ToString + " AND ENTITY_TYPE= 9 "

                    'drRow = dsSet.Tables("Documents").Select(strFilterString)
                    'noOfFiles = DocsImages.populateListView(strUserId, drRow, ActiveOwner)
                    strFilterString = String.Empty
                    drRow = Nothing
                End If


                If Not IsNothing(strEntityIds) Then
                    strListOfEntityIds = strEntityIds.Split(",")
                End If

                If Not IsNothing(strListOfEntityIds) Then
                    If UBound(strListOfEntityIds) <> -1 Then
                        For i = 0 To UBound(strListOfEntityIds)
                            If strListOfEntityIds(i).ToString <> String.Empty Then
                                nEntityId = Integer.Parse(strListOfEntityIds(i))
                                strFilterString = " ENTITY_ID=" + nEntityId.ToString + " AND ENTITY_TYPE=" + nEntityType.ToString

                                'drRow = dsSet.Tables("Documents").Select(strFilterString)
                                'noOfFiles = DocsImages.populateListView(strUserId, drRow, nEntityId)

                                strFiles += DocsImages.getfiles(nEntityId)
                            End If
                        Next
                        If strFiles.Length > 0 Then
                            noOfImages += DocsImages.images(strFiles)
                            DocsImages.ImagesOnPage()
                        End If
                    End If
                End If

            End If

            If Not noOfFiles > 0 And noOfImages > 0 Then
                DocsImages.tbCntrlDocsImages.SelectedTab = DocsImages.tbPageImages
            Else
                '  DocsImages.tbCntrlDocsImages.SelectedTab = DocsImages.tbPageDocs
            End If
            If noOfFiles <= 0 And noOfImages <= 0 Then
                MsgBox("No Documents/Images Found")
            End If

        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                Dim MyErr As ErrorReport
                MyErr = New ErrorReport(New Exception("No Documents found.", ex))
                MyErr.ShowDialog()
            Else
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            End If

        Finally
            'dConsumer = Nothing

        End Try
        If Not noOfFiles <= 0 Or Not noOfImages <= 0 Then
            DocsImages.ShowDialog()
        End If
    End Sub
    Private Sub mnuItemPrevFacs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemPrevFacs.Click

        Dim ActiveGUID As Guid
        Dim ActiveOwner As Int64
        Dim RegForm As Registration
        Dim oForm As Form

        Try
            ActiveGUID = AppSemaphores.GetValuePair("0", "ActiveForm")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                MsgBox("No windows open.  You must first open a registration window and select an owner.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Open a Registration Window")
                Exit Sub
            End If

        End Try
        Try
            ActiveOwner = MusterContainer.AppSemaphores.GetValuePair(ActiveGUID.ToString, "OWNERID", True)
            If ActiveOwner = 0 Then
                MsgBox("Please select an owner first.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Select an Owner")
                Exit Sub
            Else
                If TypeOf (ActiveMdiChild) Is Registration Then
                    CType(ActiveMdiChild, Registration).PopulatePreviouslyOwnedFacilities(CType(ActiveMdiChild, Registration).nOwnerID)
                Else
                    MsgBox("The Visible form is not a Registration Form")
                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region
#Region "Facility"
    Private Sub mnItemAddFacility_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnItemAddFacility.Click

        Dim ActiveGUID As Guid
        Dim FormGUID As Guid
        Dim ActiveOwner As Int64
        Dim ActiveOwnerName As String
        Dim msgresult As MsgBoxResult
        Dim RegForm As Registration
        Dim oForm As Form

        Try
            ActiveGUID = AppSemaphores.GetValuePair("0", "ActiveForm")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                MsgBox("No windows open.  You must first open a registration window.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Open a Registration Form")
                Exit Sub
            End If

        End Try

        ActiveOwner = AppSemaphores.GetValuePair(ActiveGUID.ToString, "OwnerID", True)
        ActiveOwnerName = AppSemaphores.GetValuePair(ActiveGUID.ToString, "OwnerName", True)

        If ActiveOwner = 0 Then
            MsgBox("Please navigate to the registration form containing the owner for which you wish to add a facility", MsgBoxStyle.OKOnly, "Registration Form Not Active")
            Exit Sub
        Else
            For Each oForm In Me.MdiChildren
                If TypeOf (oForm) Is Registration Then
                    RegForm = CType(oForm, Registration)
                    FormGUID = RegForm.MyGuid
                    If FormGUID.ToString = ActiveGUID.ToString Then
                        RegForm.nFacilityID = 0
                        RegForm.bolAddFacility = True
                        If RegForm.tbCntrlRegistration.SelectedTab.Name = RegForm.tbPageFacilityDetail.Name Then
                            RegForm.SetupTabs()
                        Else
                            RegForm.tbCntrlRegistration.SelectedTab = RegForm.tbPageFacilityDetail
                        End If
                        RegForm.Focus()
                    End If
                End If
            Next
        End If

    End Sub
    Private Sub mnuItemFacilityDocsPhotos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuItemFacilityDocsPhotos.Click
        Dim DocsImages As New DocsPhotos(Me)
        'Dim dsSet As DataSet
        Dim noOfFiles As Integer
        Dim noOfImages As Integer
        Dim nEntityId As Integer
        Dim nEntityType As Integer
        'Dim dConsumer As New DocConsumer
        Dim strFilterString As String
        Dim strActiveFormValue As String
        Dim strFiles As String = ""
        Dim i As Integer = 0
        Dim drRow() As DataRow
        Dim strFormName As String = String.Empty
        Dim TechForm As MUSTER.Technical
        Dim ClosrForm As MUSTER.Closure
        Dim FinancialForm As MUSTER.Financial
        Dim FeesForm As MUSTER.Fees
        Dim CandEForm As MUSTER.CandE
        Dim ActiveGUID As Guid
        Dim ActiveFacility As Int64
        Dim RegForm As Registration
        Dim oForm As Form

        '''''''''''''''''''''''''''''''''''
        'Hard Coded
        Dim strUserId As String = AppUser.ID
        Try
            ActiveGUID = AppSemaphores.GetValuePair("0", "ActiveForm")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                Dim MyErr As ErrorReport
                MyErr = New ErrorReport(New Exception("No Documents/Images Found.You have to select a facility", ex))
                MyErr.ShowDialog()
                Exit Sub
            End If

        End Try
        Try
            'dsSet = dConsumer.loadDocument(strUserId)
            'dsSet = pOwn.RunSQLQuery("SELECT * FROM tblSYS_DOCUMENT_MANAGER")
            ActiveFacility = MusterContainer.AppSemaphores.GetValuePair(ActiveGUID.ToString, "FacilityID", True)
            If ActiveFacility = 0 Then
                MsgBox("No Documents/Images Found.You have to select a facility")
                Exit Sub
            Else
                '    For Each oForm In Me.MdiChildren
                '        If TypeOf (oForm) Is Registration Then
                '            RegForm = CType(oForm, Registration)
                '            RegForm.tbCntrlRegistration.SelectedTab = RegForm.tbPageFacilityDetail
                '        End If
                '        If oForm.Text.StartsWith("Registration - Facility Detail") Then
                '            DocsImages.Text = oForm.Text + " Documents"
                '            nEntityType = 6
                '            strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString
                '            nEntityId = CInt(AppSemaphores.GetValuePair(strActiveFormValue, "FacilityDetails"))
                '        End If
                '    Next
                'End If

                'If oForm.Text.StartsWith("Registration - Facility Detail") Then
                Select Case cmbSearchModule.Text
                    Case "Registration"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Registration Then
                                RegForm = CType(oForm, Registration)
                                strFormName = RegForm.Name
                                RegForm.tbCntrlRegistration.SelectedTab = RegForm.tbPageFacilityDetail
                            End If
                            If oForm.Text.StartsWith("Registration - Facility Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                nEntityType = 6
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString
                                nEntityId = CInt(AppSemaphores.GetValuePair(strActiveFormValue, "FacilityDetails"))
                            End If
                        Next
                    Case "Technical"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Technical Then
                                TechForm = CType(oForm, Technical)
                                strFormName = TechForm.Name
                                TechForm.tbCntrlTechnical.SelectedTab = TechForm.tbPageFacilityDetail

                            End If
                            If oForm.Text.StartsWith(strFormName + " LUST Events - Facility") Or oForm.Text.StartsWith(strFormName + " - Facility Detail") Then
                                ' oForm.Text.StartsWith(strFormName + " - Facility Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                nEntityType = 6
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString
                                nEntityId = CInt(AppSemaphores.GetValuePair(strActiveFormValue, "FacilityDetails"))
                            End If
                        Next
                    Case "Closure"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Closure Then
                                ClosrForm = CType(oForm, Closure)
                                strFormName = ClosrForm.Name
                                ClosrForm.tbCntrlClosure.SelectedTab = ClosrForm.tbPageFacilityDetail
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Facility Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                nEntityType = 6
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString
                                nEntityId = CInt(AppSemaphores.GetValuePair(strActiveFormValue, "FacilityDetails"))
                            End If
                        Next
                    Case "Financial"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Financial Then
                                FinancialForm = CType(oForm, Financial)
                                strFormName = FinancialForm.Name
                                FinancialForm.tbCntrlFinancial.SelectedTab = FinancialForm.tbPageFacilityDetail
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Facility Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                nEntityType = 6
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString
                                nEntityId = CInt(AppSemaphores.GetValuePair(strActiveFormValue, "FacilityDetails"))
                            End If
                        Next
                    Case "Fees"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is Fees Then
                                FeesForm = CType(oForm, Fees)
                                strFormName = FeesForm.Name
                                FeesForm.tbCntrlFees.SelectedTab = FeesForm.tbPageFacilityDetail
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Facility Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                nEntityType = 6
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString
                                nEntityId = CInt(AppSemaphores.GetValuePair(strActiveFormValue, "FacilityDetails"))
                            End If
                        Next
                    Case "C & E"
                        For Each oForm In Me.MdiChildren
                            If TypeOf (oForm) Is CandE Then
                                CandEForm = CType(oForm, CandE)
                                strFormName = "C " + "&" + " E"
                                CandEForm.tbCntrlCandE.SelectedTab = CandEForm.tbPageFacilityDetail
                            End If
                            If oForm.Text.StartsWith(strFormName + " - Facility Detail") Then
                                DocsImages.Text = oForm.Text + " Documents"
                                nEntityType = 6
                                strActiveFormValue = AppSemaphores.GetValuePair("0", "ActiveForm").ToString
                                nEntityId = CInt(AppSemaphores.GetValuePair(strActiveFormValue, "FacilityDetails"))
                            End If
                        Next
                End Select
            End If



            If oForm.Text.StartsWith(strFormName + " - Facility Detail") Or oForm.Text.StartsWith(strFormName + " LUST Events - Facility") Then
                strFilterString = " ENTITY_ID=" + nEntityId.ToString + " AND ENTITY_TYPE=" + nEntityType.ToString

                'drRow = dsSet.Tables("Documents").Select(strFilterString)
                'noOfFiles = DocsImages.populateListView(strUserId, drRow, nEntityId)

                strFiles += DocsImages.getfiles(nEntityId)
                If strFiles.Length > 0 Then
                    noOfImages += DocsImages.images(strFiles)
                    DocsImages.ImagesOnPage()
                End If
            End If

            If Not noOfFiles > 0 And noOfImages > 0 Then
                DocsImages.tbCntrlDocsImages.SelectedTab = DocsImages.tbPageImages
            Else
                ' DocsImages.tbCntrlDocsImages.SelectedTab = DocsImages.tbPageDocs
            End If

            If noOfFiles <= 0 And noOfImages <= 0 Then
                MsgBox("No Documents/Images Found")
            End If

        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                Dim MyErr As ErrorReport
                MyErr = New ErrorReport(New Exception("No Documents found.", ex))
                MyErr.ShowDialog()
            Else
                Dim MyErr As ErrorReport
                MyErr = New ErrorReport(ex)
                MyErr.ShowDialog()
            End If

        Finally
            'dConsumer = Nothing

        End Try
        If Not noOfFiles <= 0 Or Not noOfImages <= 0 Then
            DocsImages.ShowDialog()
        End If

    End Sub
    Private Sub MenuItemPreOwners_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItemPreOwners.Click

        Dim ActiveGUID As Guid
        Dim ActiveFacility As Int64
        Dim RegForm As Registration
        Dim oForm As Form

        Try
            ActiveGUID = AppSemaphores.GetValuePair("0", "ActiveForm")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                MsgBox("No windows open.  You must first open a registration window and select a Facility.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Open a Registration Window")
                Exit Sub
            End If

        End Try
        Try

            ActiveFacility = MusterContainer.AppSemaphores.GetValuePair(ActiveGUID.ToString, "FacilityID", True)
            If ActiveFacility = 0 Then
                MsgBox("Please select an Facility first.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Select an Facility")
                Exit Sub
            Else
                If TypeOf (ActiveMdiChild) Is Registration Then
                    CType(ActiveMdiChild, Registration).PopulatePreviouslyOwnedOwners(CType(ActiveMdiChild, Registration).nFacilityID)
                Else
                    MsgBox("The Visible form is not a Registration Form")
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
    Private Sub mnItemTransferOwnership_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnItemTransferOwnership.Click
        Dim ActiveGUID As Guid
        Dim ActiveOwner As Int64
        Dim RegForm As Registration
        Dim oForm As Form

        Try
            ActiveGUID = AppSemaphores.GetValuePair("0", "ActiveForm")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                MsgBox("No windows open.  You must first open a registration window and select an owner.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Open a Registration Window")
                Exit Sub
            End If

        End Try
        Try
            ActiveOwner = MusterContainer.AppSemaphores.GetValuePair(ActiveGUID.ToString, "OwnerID", True)
            If ActiveOwner = 0 Then
                MsgBox("Please select an owner first.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Select an Owner")
                Exit Sub
            Else
                If TypeOf (ActiveMdiChild) Is Registration Then
                    CType(ActiveMdiChild, Registration).CallTransferOwnership()
                Else
                    MsgBox("The Visible form is not a Registration Form")
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mnuItemComplianceLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemComplianceLetter.Click
        Try
            Dim strResponse As String = InputBox("Enter Facility ID(s) (XX / XX, XX)", "Facility Compliance Letter", "")

            If strResponse <> String.Empty Then
                Dim strInvalidFacID As String = String.Empty
                Dim strFacIDs() As String
                Dim strValidFacIDs As String = String.Empty
                Dim nFacID As Integer

                strFacIDs = strResponse.Split(",")
                For Each strFacID As String In strFacIDs
                    strFacID = strFacID.Trim
                    If strFacID <> String.Empty Then
                        If Not IsNumeric(strFacID) Then
                            strInvalidFacID += strFacID + ", "
                        Else
                            nFacID = strFacID
                            If nFacID <= 0 Then
                                strInvalidFacID += strFacID + ", "
                            Else
                                strValidFacIDs = String.Format("{0}{1}{2}", strValidFacIDs, IIf(strValidFacIDs.Length > 0, ",", String.Empty), strFacID)
                            End If
                        End If
                    End If
                Next

                If strInvalidFacID = String.Empty Then
                    Dim strFiscalYear As String = InputBox("Enter Fiscal Year (xxxx)", "Facility Compliance Fiscal Year", Today.Year.ToString)
                    Dim regLetters As New Reg_Letters
                    Me.Cursor = Cursors.WaitCursor

                    regLetters.GenerateComplianceLetter(False, strValidFacIDs, strFiscalYear)
                Else
                    MsgBox("Invalid Entry - " + strInvalidFacID.TrimEnd.TrimEnd(","))
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub mnuItemNonComplianceLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemNonComplianceLetter.Click
        Try
            Dim strResponse As String = InputBox("Enter Facility ID(s) (XX / XX, XX)", "Facility Non Compliance Letter", "")

            If strResponse <> String.Empty Then
                Dim strInvalidFacID As String = String.Empty
                Dim strFacIDs() As String
                Dim alValidFacIDs As New ArrayList
                Dim nFacID As Integer

                strFacIDs = strResponse.Split(",")
                For Each strFacID As String In strFacIDs
                    strFacID = strFacID.Trim
                    If strFacID <> String.Empty Then
                        If Not IsNumeric(strFacID) Then
                            strInvalidFacID += strFacID + ", "
                        Else
                            nFacID = strFacID
                            If nFacID <= 0 Then
                                strInvalidFacID += strFacID + ", "
                            Else
                                alValidFacIDs.Add(nFacID)
                            End If
                        End If
                    End If
                Next

                If strInvalidFacID = String.Empty Then
                    Dim strFiscalYear As String = InputBox("Enter Fiscal Year (xxxx)", "Facility Compliance Fiscal Year", Today.Year.ToString)
                    Dim regLetters As New Reg_Letters
                    Me.Cursor = Cursors.WaitCursor
                    For Each nFacID In alValidFacIDs
                        regLetters.GenerateComplianceLetter(True, nFacID, strFiscalYear)
                    Next
                Else
                    MsgBox("Invalid Entry - " + strInvalidFacID.TrimEnd.TrimEnd(","))
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub mnuTOSILetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTOSILetter.Click
        Try
            Dim strResponse As String = InputBox("Enter Facility ID(s) (XX / XX, XX)", "TOS-I Letter", "")

            If strResponse <> String.Empty Then
                Dim strInvalidFacID As String = String.Empty
                Dim strFacIDs() As String
                Dim strValidFacIDs As String = String.Empty
                Dim nFacID As Integer

                strFacIDs = strResponse.Split(",")
                For Each strFacID As String In strFacIDs
                    strFacID = strFacID.Trim
                    If strFacID <> String.Empty Then
                        If Not IsNumeric(strFacID) Then
                            strInvalidFacID += strFacID + ", "
                        Else
                            nFacID = strFacID
                            If nFacID <= 0 Then
                                strInvalidFacID += strFacID + ", "
                            Else
                                strValidFacIDs = String.Format("{0}{1}{2}", strValidFacIDs, IIf(strValidFacIDs.Length > 0, ",", String.Empty), strFacID)
                            End If
                        End If
                    End If
                Next

                If strInvalidFacID = String.Empty Then
                    '  Dim strFiscalYear As String = InputBox("Enter Fiscal Year (xxxx)", "Facility Compliance Fiscal Year", Today.Year.ToString)
                    Dim regLetters As New Reg_Letters
                    Me.Cursor = Cursors.WaitCursor

                    regLetters.GenerateTOSILetter(False, strValidFacIDs)
                Else
                    MsgBox("Invalid Entry - " + strInvalidFacID.TrimEnd.TrimEnd(","))
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
#End Region
#Region "CAP"
    Private Sub mnuCAPPreMonthly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCAPPreMonthly.Click
        Dim objCAPPreMonthly As CAPPreMonthly

        Dim localGUID As String
        Dim ChildForm As Windows.Forms.Form
        Dim fGUID As System.Guid
        Try
            If localGUID Is Nothing Then localGUID = String.Empty
            localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("WindowName", "CAP PreMonthly", "Registration")
            If Not localGUID = String.Empty Then  'Existing inspection
                'Existing Inspection Window
                For Each ChildForm In Me.MdiChildren
                    If ChildForm.GetType.Name = "Registration" Then
                        fGUID = CType(ChildForm, CAPPreMonthly).MyGuid
                        If fGUID.ToString = localGUID Then
                            ChildForm.Activate()
                            objCAPPreMonthly = CType(ChildForm, CAPPreMonthly)
                            FillRegForm(0, 0, "", , , , , , , , , , objCAPPreMonthly)
                            Exit Sub
                        End If
                    End If
                Next
                MusterContainer.AppSemaphores.Remove(localGUID)
                mnuCAPPreMonthly_Click(sender, e)
            Else
                'New CAPPreMonthly
                objCAPPreMonthly = New CAPPreMonthly
                objCAPPreMonthly.MdiParent = Me
                objCAPPreMonthly.Show()
                objCAPPreMonthly.BringToFront()
                objCAPPreMonthly.Tag = "CAP PreMonthly"
                HideShowOwnerInfo(False)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub mnuCAPMonthly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCAPMonthly.Click

        Dim cap As CAP_Letters
        Try
            cap = New CAP_Letters
            cap.SetupSystemToGenerateCAPMonthly(True, String.Empty, -1)
        Catch ex As Exception

            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            cap = Nothing
        End Try
    End Sub


    Private Sub mnuSpecialCAPMonthly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSpecialCAPMonthly.Click
        Try
            'check rights for Tank & Pipe
            If Not AppUser.HasAccess(CType(UIUtilsGen.ModuleID.CAPProcess, Integer), MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Tank) Then
                MessageBox.Show("You do not have Rights to Monthly CAP Processing")
                Exit Sub
            ElseIf Not AppUser.HasAccess(CType(UIUtilsGen.ModuleID.CAPProcess, Integer), MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Pipe) Then
                MessageBox.Show("You do not have Rights to Monthly CAP Processing")
                Exit Sub
            End If

            Dim strMonthYear As String
            Dim ownerName As String = String.Empty
            Dim processingDateMonth As Date
            strMonthYear = InputBox("Enter the Month (XX) / Year (XXXX)", "Monthly CAP Processing", Today.Month.ToString + "/" + Today.Year.ToString)
            If strMonthYear = String.Empty Then
                Exit Sub
            ElseIf strMonthYear.Split("/").Length < 2 Then
                MsgBox("Invalid entry")
                Exit Sub
            End If

            Try
                processingDateMonth = strMonthYear.Split("/")(0) + "/1/" + strMonthYear.Split("/")(1)
            Catch ex As Exception
                MsgBox("Invalid entry")
                Exit Sub
            End Try

            ownerName = InputBox("Enter Owner Name (Leave Blank for Full Monthly CAP Report)", , String.Empty)


            Dim capLetters As New CAP_Letters
            Me.Cursor = Cursors.WaitCursor
            capLetters.GenerateCAPMonthlyForAll(processingDateMonth, pOwn, ownerName)
            'Dim CAPcons As New CAP_Letters
            'Me.Cursor = Cursors.WaitCursor
            'CAPcons.getOwnerDetailsForCAPProcessing(processingDateMonth)
            'Me.Cursor = Cursors.Default
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    'Private Sub mnuCALetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCALetter.Click
    '    Try
    '        Dim strMonthYear As String
    '        strMonthYear = InputBox("Enter - Month / Year for Monthly" + vbCrLf + vbTab + "OR" + vbCrLf + "Enter - Year for Yearly", "Compliance Assistance Letter", Today.Month.ToString + "/" + Today.Year.ToString)
    '        Dim yr, mnth As Long
    '        If strMonthYear = String.Empty Then
    '            Exit Sub
    '        End If
    '        Try
    '            If strMonthYear.Split("/").Length > 1 Then
    '                mnth = strMonthYear.Split("/")(0)
    '                yr = strMonthYear.Split("/")(1)
    '            Else
    '                yr = strMonthYear
    '            End If
    '        Catch ex As Exception
    '            Throw ex
    '        End Try
    '        Dim CAPcons As New CAP_Letters
    '        Me.Cursor = Cursors.WaitCursor
    '        CAPcons.GenericCAPLetterGenerationforAllOwners("CAP", yr, mnth)
    '        Me.Cursor = Cursors.Default
    '    Catch ex As Exception
    '        Me.Cursor = Cursors.Default
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub mnuCAPYearly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCAPYearly.Click

        Dim cap As CAP_Letters
        Try
            cap = New CAP_Letters
            cap.SetupSystemToGenerateCAPYearly(CAP_Letters.CapAnnualMode.StaticByYear, True, String.Empty, -1)
        Catch ex As Exception

            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            cap = Nothing
        End Try
    End Sub


    Private Sub mnuCAPCurrentSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCAPCurrent.Click

        Dim cap As CAP_Letters
        Try
            cap = New CAP_Letters
            cap.SetupSystemToGenerateCAPYearly(CAP_Letters.CapAnnualMode.CurrentSummary, True, Nothing, -1)
        Catch ex As Exception

            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            cap = Nothing
        End Try

    End Sub
#End Region
#Region "Fees"
    Private Sub mnuSubItemAdminServices_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSubItemAdminServices.Click
        Dim objAdminServices As AdministrativeServices
        Try
            objAdminServices = New AdministrativeServices
            objAdminServices.MdiParent = Me
            objAdminServices.Show()
            objAdminServices.BringToFront()
            objAdminServices.Tag = "AdministrativeServices"
            HideShowOwnerInfo(True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub mnuSubItemFees_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSubItemFees.Click
    '    Dim objFees As Fees
    '    Try
    '        If pOwn.ID > 0 Then
    '            objFees = New Fees(pOwn)
    '            objFees.MdiParent = Me
    '            objFees.Show()
    '            objFees.BringToFront()
    '            objFees.Tag = "Fees"
    '            HideShowOwnerInfo(True)
    '        Else
    '            MessageBox.Show("Please select Owner First")
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region
#Region "Company"
#Region "Company"
    Private Sub mnuItemAddCompany_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemAddCompany.Click
        Dim objComp As Company
        Try
            objComp = New Company(0)
            objComp.MdiParent = Me
            objComp.Show()
            objComp.BringToFront()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Licensees"
    Private Sub mnuItemAddLicensee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemAddLicensee.Click
        Dim objLic As Licensees
        Try
            objLic = New Licensees(, , , "ADD")
            objLic.MdiParent = Me
            objLic.Show()
            objLic.WindowState = FormWindowState.Maximized
            objLic.BringToFront()
            objLic.pnlCompany.Visible = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub mnuItemAddComplianceManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemAddComplianceManager.Click
        Dim objMgr As Managers
        Try
            objMgr = New Managers(, , , "ADD")
            objMgr.MdiParent = Me
            objMgr.Show()
            objMgr.WindowState = FormWindowState.Maximized
            objMgr.BringToFront()
            objMgr.pnlCompany.Visible = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mnuItemLicenseeMgmt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemLicenseeMgmt.Click
        Dim objLM As LicenseeManagement
        Try
            objLM = New LicenseeManagement
            objLM.MdiParent = Me
            objLM.Show()
            objLM.BringToFront()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region
#Region "Technical"
    'Private Sub mnItemTechnical_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnItemTechnical.Click
    '    Dim objTechnical As Technical
    '    Try
    '        If pOwn.ID > 0 Then
    '            objTechnical = New Technical(pOwn)
    '            objTechnical.MdiParent = Me
    '            objTechnical.Show()
    '            objTechnical.BringToFront()
    '            objTechnical.Tag = "Technical"
    '            HideShowOwnerInfo(True)
    '            objTechnical.btnViewModifyComment.Enabled = False
    '        Else
    '            MessageBox.Show("Please select Owner First.")
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub mnuItemRemSysHist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemRemSysHist.Click
        Dim objRemSysList As RemediationSystemList
        Try
            objRemSysList = New RemediationSystemList
            objRemSysList.MdiParent = Me
            objRemSysList.Mode = 1
            objRemSysList.EventActivityID = 0
            objRemSysList.Show()
            objRemSysList.BringToFront()
            objRemSysList.Tag = "RemediationSystem"
            HideShowOwnerInfo(True)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "C&E"
    Private Sub mnuItemCAEManagement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemCAEManagement.Click
        Dim objCAEMgmt As CandEManagement

        Dim localGUID As String
        Dim ChildForm As Windows.Forms.Form
        Dim fGUID As System.Guid
        Try
            If localGUID Is Nothing Then localGUID = String.Empty
            localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("WindowName", "C and E Management", "CandEManagement")
            If Not localGUID = String.Empty Then  'Existing inspection
                'Existing Inspection Window
                For Each ChildForm In Me.MdiChildren
                    If ChildForm.GetType.Name = "CandEManagement" Then
                        fGUID = CType(ChildForm, CandEManagement).MyGuid
                        If fGUID.ToString = localGUID Then
                            ChildForm.Activate()
                            objCAEMgmt = CType(ChildForm, CandEManagement)
                            FillRegForm(0, 0, "", , , , , , , , objCAEMgmt)
                            Exit Sub
                        End If
                    End If
                Next
                MusterContainer.AppSemaphores.Remove(localGUID)
                mnuItemCAEManagement_Click(sender, e)
            Else
                'New Inspection
                objCAEMgmt = New CandEManagement
                objCAEMgmt.MdiParent = Me
                objCAEMgmt.Show()
                objCAEMgmt.BringToFront()
                objCAEMgmt.Tag = "C and E Management"
                HideShowOwnerInfo(False)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mnuItemCAEInspectors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Dim objCAEIns As Inspectors

        Dim localGUID As String
        Dim ChildForm As Windows.Forms.Form
        Dim fGUID As System.Guid
        Try
            If localGUID Is Nothing Then localGUID = String.Empty
            localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("WindowName", "C and E Inspectors", "CandEManagement")
            If Not localGUID = String.Empty Then  'Existing inspection
                'Existing Inspection Window
                For Each ChildForm In Me.MdiChildren
                    If ChildForm.GetType.Name = "Inspectors" Then
                        fGUID = CType(ChildForm, Inspectors).MyGuid
                        If fGUID.ToString = localGUID Then
                            ChildForm.Activate()
                            objCAEIns = CType(ChildForm, Inspectors)
                            FillRegForm(0, 0, "", , , , , , , , , objCAEIns)
                            Exit Sub
                        End If
                    End If
                Next
                MusterContainer.AppSemaphores.Remove(localGUID)
                mnuItemCAEInspectors_Click(sender, e)
            Else
                'New C&E Inspectors
                objCAEIns = New Inspectors
                objCAEIns.MdiParent = Me
                objCAEIns.Show()
                objCAEIns.BringToFront()
                objCAEIns.Tag = "C and E Inspectors"
                HideShowOwnerInfo(False)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
    Private Sub mnuItemInspector_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuItemInspector.Click
        Dim objInspection As Inspection

        Dim localGUID As String
        Dim ChildForm As Windows.Forms.Form
        Dim fGUID As System.Guid
        Try
            If localGUID Is Nothing Then localGUID = String.Empty
            localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("WindowName", "Inspection Schedule", "Inspection")
            If Not localGUID = String.Empty Then  'Existing inspection
                'Existing Inspection Window
                For Each ChildForm In Me.MdiChildren
                    If ChildForm.GetType.Name = "Inspection" Then
                        fGUID = CType(ChildForm, Inspection).MyGuid
                        If fGUID.ToString = localGUID Then
                            ChildForm.Activate()
                            objInspection = CType(ChildForm, Inspection)
                            FillRegForm(0, 0, "", , , , , , objInspection)
                            Exit Sub
                        End If
                    End If
                Next
                MusterContainer.AppSemaphores.Remove(localGUID)
                mnuItemInspector_Click(sender, e)
            Else
                'New Inspection
                objInspection = New Inspection(pInspec)
                objInspection.MdiParent = Me
                objInspection.Show()
                objInspection.BringToFront()
                objInspection.Tag = "Inspection"
                HideShowOwnerInfo(False)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#Region "Closure"
    'Private Sub mnuItemClosure_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuItemClosure.Click
    '    Dim objClosure As Closure
    '    Try
    '        If pOwn.ID > 0 Then
    '            objClosure = New Closure(pOwn)
    '            objClosure.MdiParent = Me
    '            objClosure.Show()
    '            objClosure.BringToFront()
    '            objClosure.Tag = "Closure"
    '            HideShowOwnerInfo(True)
    '        Else
    '            MessageBox.Show("Please select Owner First.")
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub mnuCancelAllOverDue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCancelAllOverDue.Click
        Dim ActiveGUID As Guid
        Dim ActiveFacility As Int64
        Dim colsureForm As Closure

        Try
            ActiveGUID = AppSemaphores.GetValuePair("0", "ActiveForm")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                MsgBox("No windows open.  You must first open a Closure window and select a Facility.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Open a Closure Window")
                Exit Sub
            End If

        End Try
        Try
            ActiveFacility = MusterContainer.AppSemaphores.GetValuePair(ActiveGUID.ToString, "FacilityID", True)
            If ActiveFacility = 0 Then
                MsgBox("Please select an Facility first.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Please Select an Facility")
                Exit Sub
            Else
                If TypeOf (ActiveMdiChild) Is Closure Then
                    CType(ActiveMdiChild, Closure).CancelAllOverdue()
                Else
                    MsgBox("The Visible form is not a Closure Form")
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
#End Region
#Region "Contact"
    Private Sub mnuContactReconciliation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuContactReconciliation.Click

        Return

        Dim objReconcile As ContactReconciliation
        Try
            objReconcile = New ContactReconciliation(pConStruct)
            objReconcile.MdiParent = Me
            objReconcile.Show()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Financial"
    'Private Sub mnuItemFinancial_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuItemFinancial.Click
    '    Dim objFinancial As Financial
    '    Try
    '        If pOwn.ID > 0 Then
    '            objFinancial = New Financial(pOwn)
    '            objFinancial.MdiParent = Me
    '            objFinancial.Show()
    '            objFinancial.BringToFront()
    '            objFinancial.Tag = "Financial"
    '            HideShowOwnerInfo(True)
    '        Else
    '            MessageBox.Show("Please select Owner First.")
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub mnuFinRollover_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFinRollover.Click
        Dim oRollovers As New Rollovers
        oRollovers.MdiParent = Me
        oRollovers.Mode = 1
        oRollovers.Show()
    End Sub
    Private Sub mnuFinRolloverPO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFinRolloverPO.Click
        Dim oRollovers As New Rollovers
        oRollovers.MdiParent = Me
        oRollovers.Mode = 2
        oRollovers.Show()
    End Sub
#End Region
    Private Sub mnuItemExitApp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemExitApp.Click
        Me.Close()
    End Sub
#End Region
#Region "Utilities"
    Private Sub mnItemRegReports_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnItemRegReports.Click
        Dim frmReport As ReportDisplay

        frmReport = New ReportDisplay
        frmReport.MdiParent = Me
        Try
            frmReport.Show()
            If Not frmReport.cboModule.DataSource Is Nothing Then
                frmReport.cboModule.SelectedIndex = -1
                UIUtilsGen.SetComboboxItemByValue(frmReport.cboModule, AppUser.DefaultModule)
            End If

            'frmReport.cboModule.SelectedIndex = 0
            'frmReport.cboModule.SelectedIndex = -1
            'frmReport.ComboBox1.SelectedIndex = -1
            'frmReport.ComboBox1.SelectedIndex = -1
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error in loading reports form : " & vbCrLf & ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub mnuLetters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLetters.Click
        Dim frmLtr As Letters

        frmLtr = New Letters
        frmLtr.MdiParent = Me
        Try
            frmLtr.WindowState = FormWindowState.Maximized
            frmLtr.BringToFront()
            frmLtr.Show()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Displaying Letters form : " & vbCrLf & ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mnuItemSyncManualDocs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemSyncManualDocs.Click
        Try
            Dim isAdmin As Boolean = False
            Dim bolAllUser As Boolean = False
            Dim DOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_ManuallyCreated).ProfileValue & "\"
            Dim DOC_PATH_MAIN As String = DOC_PATH
            If DOC_PATH = "\" Then
                MsgBox("Document Path Unspecified. Please give the path before generating the letter.")
                Exit Sub
            End If

            For Each mnuItem As MenuItem In mnuMain.MenuItems
                If mnuItem.Text = "&Admin" Then
                    isAdmin = True
                    Exit For
                End If
            Next

            If isAdmin Then
                Dim result As DialogResult
                result = MessageBox.Show("Do you want to Sync all users Documents?", "Sync Manual Documents", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If result = DialogResult.Yes Then
                    bolAllUser = True
                ElseIf result = DialogResult.Cancel Then
                    Exit Sub
                End If
            End If

            Dim oLetter As New MUSTER.BusinessLogic.pLetter
            'Get the Manual Documents with modified Description.
            Dim dsDocs As DataSet
            dsDocs = oLetter.GetManualDocsWithDesc(IIf(bolAllUser, String.Empty, MusterContainer.AppUser.ID))

            Dim arrDirectories As New ArrayList
            Dim fileinfo As System.IO.FileInfo
            'Clear all the database entries for the current user.
            oLetter.DeleteManualDocuments(IIf(bolAllUser, String.Empty, MusterContainer.AppUser.ID))

            If Not bolAllUser Then
                DOC_PATH = DOC_PATH & MusterContainer.AppUser.ID & "\"
            End If

            If System.IO.Directory.Exists(DOC_PATH) Then
                'get files from the source directory
                GetFileNamesFromDirectory(DOC_PATH, arrDirectories)
                'get files from the sub-directory
                GetRecursiveFiles(DOC_PATH, arrDirectories)
                If arrDirectories.Count > 0 Then
                    Dim enDirectory As System.Collections.IEnumerator
                    Dim strOwningUser As String = String.Empty
                    Dim i As Integer = 0
                    Dim index As Integer = 0
                    Dim strModuleName As String = String.Empty
                    Dim nModuleID As Integer = 0
                    enDirectory = arrDirectories.GetEnumerator

                    While enDirectory.MoveNext
                        fileinfo = CType(enDirectory.Current, System.IO.FileInfo)

                        strOwningUser = String.Empty
                        i = 0
                        index = 0
                        strModuleName = String.Empty
                        nModuleID = 0

                        For i = 0 To fileinfo.Name.Length - 1
                            If Not System.Char.IsNumber(fileinfo.Name.Chars(i)) Then
                                index = i
                                Exit For
                            End If
                        Next
                        Dim strentityID As String = fileinfo.Name.Substring(0, index)
                        strentityID = IIf(strentityID = String.Empty, "0", strentityID)
                        Dim nentityID As Integer
                        If IsNumeric(strentityID) Then
                            nentityID = strentityID
                        Else
                            nentityID = 0
                        End If


                        Dim strUserAndModule As String = fileinfo.FullName.Substring(DOC_PATH_MAIN.Length, fileinfo.FullName.Length - DOC_PATH_MAIN.Length)
                        strOwningUser = strUserAndModule.Split("\")(0)
                        strModuleName = strUserAndModule.Split("\")(1)
                        If Not strModuleName = String.Empty Then
                            nModuleID = GetModuleIDForName(strModuleName)
                        End If

                        'Update Description for the matching entity.
                        Dim docLocation As String
                        Dim strDocDesc As String = "Manually Created Document"
                        Dim dtDocs As DataTable = dsDocs.Tables(0)
                        Dim j As Integer = 0
                        For j = 0 To dtDocs.Rows.Count - 1
                            If fileinfo.Name = dtDocs.Rows(j).Item("Document_Name").ToString() And _
                                    nentityID = dtDocs.Rows(j).Item("Entity_Id") And _
                                    strOwningUser = dtDocs.Rows(j).Item("Owning_User") Then
                                strDocDesc = dtDocs.Rows(j).Item("Document_Description")
                            End If
                        Next
                        If Not fileinfo.DirectoryName.EndsWith("\") Then
                            'fileinfo.DirectoryName.Insert(fileinfo.DirectoryName.Length(), "\\")
                            docLocation = fileinfo.DirectoryName + "\"
                        End If
                        'save file info into database.
                        Dim ltrInfo As New MUSTER.Info.LetterInfo(0, _
                                                                fileinfo.Name.Trim, _
                                                                "Manual", _
                                                                docLocation, _
                                                                6, _
                                                                nentityID, _
                                                                strDocDesc, _
                                                                1, _
                                                                CDate("01/01/0001"), _
                                                                False, _
                                                                MusterContainer.AppUser.ID, _
                                                                fileinfo.LastWriteTime, _
                                                                String.Empty, _
                                                                fileinfo.LastWriteTime, _
                                                                strOwningUser, _
                                                                nModuleID, 0, 0, 0)
                        oLetter.Save(ltrInfo, "MANUAL")
                    End While
                    MessageBox.Show("Manual Documents for " + IIf(bolAllUser, "All users", MusterContainer.AppUser.ID) + " has been Synchronized successfully.")
                Else
                    MessageBox.Show("No documents whose name starting with numbers were found!")
                End If

            Else
                MsgBox("Directory does not exists: " + DOC_PATH)
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mnuItemOpenFacPicsFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuItemOpenFacPicsFolder.Click
        Try
            Dim procStart As New ProcessStartInfo("explorer")
            Dim winDir As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_FacImages).ProfileValue
            procStart.Arguments = winDir
            Process.Start(procStart)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Help"
    Private Sub mnItemAboutMuster_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnItemAboutMuster.Click
        Dim MyFrm As New MusterAbout(Me)
        Try
            MyFrm.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mnItemCloseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnItemCloseAll.Click
        Try





            For Each frm As Form In Me.MdiChildren

                If Not pOwn Is Nothing AndAlso pOwn.colIsDirty AndAlso Me.DirtyIgnored = -1 Then
                    Me.DirtyIgnored = MsgBox("Data has been changed. Do you wish To Save changes", MsgBoxStyle.YesNo)
                End If

                frm.Close()
            Next

            DirtyIgnored = -1

            GridMaster.GlobalInstance.CleanEntityDictionary()


            While MnHistoryItems.MenuItems.Count > 0
                MnHistoryItems.MenuItems.RemoveAt(0)
            End While

            While mnItemWindows.MenuItems.Count > 0
                mnItemWindows.MenuItems.RemoveAt(0)
            End While

            btnCloseForm.Enabled = False
            btnCloseAll.Enabled = False
            mnItemCloseAll.Enabled = False
            ClearBarometer()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub StartUnzipBackupFile(ByVal strDirPath As String, ByVal strZipFileName As String)

        Dim strSQLPath As String
        Dim dirInfo As IO.DirectoryInfo

        strSQLPath = "C:\Program Files (x86)\Microsoft SQL Server\MSSQL"
        '   Try
        '  strSQLPath = LocalUserSettings.LocalMachine.OpenSubKey("SOFTWARE", False).OpenSubKey("Microsoft", False).OpenSubKey("MSSQLServer", False).OpenSubKey("Setup", False).GetValue("SQLPath")
        ' Catch ex As Exception
        '    MsgBox("Could not detect MSDE Installation on Computer. Please contact System Administrator", MsgBoxStyle.Critical)
        '   Exit Sub
        '  End Try

        ReDim aunzipParameters(3)


        aunzipParameters(0) = strDirPath + IO.Path.DirectorySeparatorChar + strZipFileName
        aunzipParameters(1) = strDirPath
        aunzipParameters(2) = 10000000000000

        Dim test As New ICSharpCode.SharpZipLib.Zip.ZipFile(aunzipParameters(0))
        aunzipParameters(2) = test.EntryByIndex(0).Size().ToString
        aunzipParameters(3) = String.Format("{0}\{1}", aunzipParameters(1), test.EntryByIndex(0).Name)

        test = Nothing

        Dim code(2) As String
        code(0) = "PrepareUnzip"
        code(1) = aunzipParameters(2)
        code(2) = aunzipParameters(3)


        RaiseEvent StartProgressScreen("Database Synchronization - Unzipping backup File", Convert.ToInt64(aunzipParameters(2)) / 100000, 0, "Preparing Backup File to be unzipped. Please wait...", code)


    End Sub

    Private Sub ContinueUnzipbackupFile()



        Dim fastZip As ICSharpCode.SharpZipLib.Zip.FastZip

        fastZip = New ICSharpCode.SharpZipLib.Zip.FastZip

        fastZip.ExtractZip(aunzipParameters(0), aunzipParameters(1), "")

        fastZip = Nothing

        Dim code(2) As String
        code(0) = "CompleteUnzip"
        code(1) = 0
        code(2) = 0

        RaiseEvent FireProgressMessage("unzipping Completed", Convert.ToInt64(aunzipParameters(2)) / 100000, code(0))

    End Sub

    Private Sub ProgressScreenEventFired(ByVal e As String) Handles ProgressScreen.ReturningMessage

        If Not e Is Nothing AndAlso e.Length > 0 Then

            Dim strDirPath As String = System.Windows.Forms.Application.StartupPath() + _
                                       IIf(System.Windows.Forms.Application.StartupPath().EndsWith(IO.Path.DirectorySeparatorChar), "", IO.Path.DirectorySeparatorChar) + "LocalDB"
            Dim strFileName As String = "Muster_Prd_db_" + Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString + "*.BAK"

            Try

                Select Case e

                    Case "ReadyForUnzip"

                        UnzipThread = New System.Threading.Thread(AddressOf ContinueUnzipbackupFile)
                        UnzipThread.Start()

                    Case "UnzipCompletedAcknowledged"
                        UnzipThread = Nothing
                        System.Threading.Thread.Sleep(500)
                        RaiseEvent FireCloseProgressScreen()
                        System.Threading.Thread.Sleep(300)
                        PerformRestoreAfterUnZip(strDirPath + IO.Path.DirectorySeparatorChar + Dir(strDirPath + IO.Path.DirectorySeparatorChar + strFileName))

                    Case "RestoreCompletedAcknowledged"
                        RestoreThread = Nothing
                        RaiseEvent FireCloseProgressScreen()
                        Dim msg As String = String.Empty


                        Try

                            If IO.File.Exists(String.Format("{0}{1}{2}", strDirPath, IO.Path.DirectorySeparatorChar, strFileName)) Then
                                System.IO.File.Delete(String.Format("{0}{1}{2}", strDirPath, IO.Path.DirectorySeparatorChar, strFileName))
                            End If

                            If IO.File.Exists(String.Format("{0}{1}{2}", strDirPath, IO.Path.DirectorySeparatorChar, strFileName.Replace(".BAK", ".ZIP"))) Then
                                System.IO.File.Delete(String.Format("{0}{1}{2}", strDirPath, IO.Path.DirectorySeparatorChar, strFileName.Replace(".BAK", ".ZIP")))
                            End If
                        Catch ex As Exception
                            msg = "Zip and Backup File of Database has not been Successfully Deleted."

                        End Try


                        MsgBox(String.Format("Database Synchronized Successfully!{0}", msg))

                End Select

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End If

    End Sub



    Private Sub DoRestore()

        Try

            Dim code(0) As String
            code(0) = "PrepareRestore"

            RaiseEvent StartProgressScreen("Database Synchronization - Data Restore", 100, 0, "Preparing Backup File to be Restored. Please wait...", code)

            Dim DbFile As String = aRestoreParameters(1)
            Dim LogFile As String = aRestoreParameters(2)

            oSQLServer = New SQLDMO.SQLServer
            oRestore = New SQLDMO.RestoreClass



            With oSQLServer

                .Connect(System.Net.Dns.GetHostName, "sa", "password")

            End With

            With oRestore

                .FileNumber = 1

                .ReplaceDatabase = True
                .Action = SQLDMO.SQLDMO_RESTORE_TYPE.SQLDMORestore_Database
                .RelocateFiles = String.Format("[Muster_Prd_Data],[{0}],[Muster_Prd_Log],[{1}] ", DbFile, LogFile)

                '.Devices = "FILES"

                .Files = String.Format("[{0}]", aRestoreParameters(0))

                .Database = "Muster_Prd"


                .ReplaceDatabase = True

                AddHandler .PercentComplete, AddressOf UpdateProgress

                .SQLRestore(oSQLServer)


            End With


            RaiseEvent FireProgressMessage("Synchronization Data Restore Successfully Executed. Finishing Up Details", 100, String.Empty)

            FinishUpSync()

            oRestore = Nothing
            oSQLServer.DisConnect()
            oSQLServer = Nothing

        Catch ex As Exception
            MsgBox(String.Format("Restoring Thread Error: {0}", ex.ToString), , "DB Restore Thread")
        End Try




    End Sub



    Private Sub UpdateProgress(ByVal message As String, ByVal percent As Integer)

        RaiseEvent FireProgressMessage(String.Format("Restoring Database at {0} percent", percent), percent, String.Empty)

    End Sub

    Private Sub FinishUpSync()

        Dim strSQL As String
        Dim strDirPath As String = System.Windows.Forms.Application.StartupPath() + _
                                   IIf(System.Windows.Forms.Application.StartupPath().EndsWith(IO.Path.DirectorySeparatorChar), "", IO.Path.DirectorySeparatorChar) + "LocalDB"
        Dim strFileName As String = "Muster_Prd_db_" + Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString + "*.BAK"
        Dim strZipFileName As String = "Muster_Prd_db_" + Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString + "*.ZIP"
        Dim strConnStr As String = "Data Source=" + System.Net.Dns.GetHostName + ";Initial Catalog=master;user='sa';password='password';"

        strConnStr = strConnStr.Substring(0, strConnStr.Length)
        LocalUserSettings.CurrentUser.SetValue("MusterSQLConnection", strConnStr)


        Try
            ' change the file paths in the local db to point to local db
            strSQL = "use muster_prd; update tblsys_profile_info set profile_value = replace(profile_value, '\\gard-prod\', '\\" + System.Net.Dns.GetHostName + "\') " + _
                        "where profile_key = 'common_paths' and profile_value like '\\gard-prod\%'"

            SqlHelper.ExecuteNonQuery(strConnStr, CommandType.Text, strSQL)

        Catch ex As SqlClient.SqlException

            MsgBox("Error: " + ex.ToString, , "SQL Error")

        End Try

        RestoreThread = Nothing

        Dim delFiles = Dir(strDirPath + IO.Path.DirectorySeparatorChar + "*.BAK")
        '   Do
        If delFiles <> String.Empty Then
            IO.File.Delete(strDirPath + IO.Path.DirectorySeparatorChar + delFiles)
            delFiles = Dir()
        End If

        '  Loop Until delFiles = String.Empty

        delFiles = Dir(strDirPath + IO.Path.DirectorySeparatorChar + "*.ZIP")
        'Do
        If delFiles <> String.Empty Then
            IO.File.Delete(strDirPath + IO.Path.DirectorySeparatorChar + delFiles)
            delFiles = Dir()
        End If

        'Loop Until delFiles = String.Empty


        RaiseEvent FireProgressMessage("Synchronization Completed", 100, "RestoreCompleted")

    End Sub



    Private Sub PerformRestoreAfterUnZip(ByVal strFileName As String)

        'Dim strDirPath As String = System.Windows.Forms.Application.StartupPath() + _
        '                           IIf(System.Windows.Forms.Application.StartupPath().EndsWith(IO.Path.DirectorySeparatorChar), "", IO.Path.DirectorySeparatorChar) + "LocalDB"
        'Dim strFileName As String = "Muster_Prd_db_" + Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString + "*.BAK"
        ' Dim strZipFileName As String = "Muster_Prd_db_" + Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString + "*.ZIP"
        Dim strServerShare As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_DBSync).ProfileValue

        Dim strConnStrPrev, strSQLPath As String

        strSQLPath = "C:\Program Files (x86)\Microsoft SQL Server\MSSQL"
        'strSQLPath = "C:\Program Files\Microsoft SQL Server\MSSQL"
        'get SQL Path
        '  Try
        '  strSQLPath = LocalUserSettings.LocalMachine.OpenSubKey("SOFTWARE", False).OpenSubKey("Microsoft", False).OpenSubKey("MSSQLServer", False).OpenSubKey("Setup", False).GetValue("SQLPath")
        '   Catch ex As Exception
        '       MsgBox("Could not detect MSDE Installation on Computer. Please contact System Administrator", MsgBoxStyle.Critical)
        '   Exit Sub
        ' End Try


        ' change registry entry to point to local db
        Dim strConnStr As String = "Data Source=" + System.Net.Dns.GetHostName + ";Initial Catalog=master;user='sa';password='password';"

        strConnStr = strConnStr.Substring(0, strConnStr.Length)
        LocalUserSettings.CurrentUser.SetValue("MusterSQLConnection", strConnStr)

        ' create db
        Try


            Dim strSQL As String = "" + _
                                   "IF EXISTS(SELECT * FROM SYSDATABASES WHERE NAME LIKE '%Muster_Prd%') " + _
                                    "BEGIN DROP DATABASE Muster_Prd  END;"



            Dim strSQL2 As String = strSQL + _
                                   "IF NOT EXISTS(SELECT * FROM SYSDATABASES WHERE NAME LIKE '%Muster_Prd%') " + _
                                    "BEGIN CREATE DATABASE Muster_Prd ON (NAME = Muster_Prd_Data, FILENAME = '" + strSQLPath + "\data\Muster_Prd_Data.MDF' ) " + _
                                    "LOG ON ( NAME = Muster_Prd_Log, FILENAME = '" + strSQLPath + "\data\Muster_Prd_Log.LDF' )" + _
                                     "  END; "

            SqlHelper.ExecuteNonQuery(strConnStr, CommandType.Text, strSQL2)



            System.Threading.Thread.Sleep(4000)

            RestoreThread = New Threading.Thread(AddressOf DoRestore)

            ReDim aRestoreParameters(2)
            aRestoreParameters(0) = strFileName
            aRestoreParameters(1) = strSQLPath + "\data\Muster_Prd_Data.MDF"
            aRestoreParameters(2) = strSQLPath + "\data\Muster_Prd_Log.LDF"


            RestoreThread.Start()

        Catch ex As SqlClient.SqlException

            MsgBox("Error: " + ex.ToString, , "SQL Error")

        Catch ex As Exception

            MsgBox("general Error: " + ex.ToString, , "General error")

        End Try


    End Sub

    Private Sub mnItemSyncDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnItemSyncDB.Click
        If MsgBox("Are you sure you want to Synchronize the Database?", MsgBoxStyle.YesNo, "Confirm Action") = MsgBoxResult.No Then
            Exit Sub
        End If



        Dim strDirPath As String = System.Windows.Forms.Application.StartupPath() + _
                                   IIf(System.Windows.Forms.Application.StartupPath().EndsWith(IO.Path.DirectorySeparatorChar), "", IO.Path.DirectorySeparatorChar) + "LocalDB"

        Dim strFileName As String = "Muster_Prd_db_" + Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString + "*.BAK"
        Dim strZipFileName As String = "Muster_Prd_db_" + Now.Year.ToString + IIf(Now.Month < 10, "0", "") + Now.Month.ToString + IIf(Now.Day < 10, "0", "") + Now.Day.ToString + "*.ZIP"
        Dim strServerShare As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_DBSync).ProfileValue
        Dim bolZipExists, bolFileExists As Boolean
        Dim strConnStrPrev As String
        Dim dirInfo As IO.DirectoryInfo
        Dim thiszipfilename As String = String.Empty

        Try
            strConnStrPrev = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")
        Catch ex As Exception
            strConnStrPrev = ""
        End Try

        '  MsgBox("strDirPath = " + strDirPath)
        Try
            Me.Cursor = Cursors.AppStarting

            bolZipExists = False
            bolFileExists = False

            ' if server path is not valid, do not proceed
            If strServerShare = String.Empty Then
                MsgBox("Invalid Path. Please contact System Administrator")
                Exit Sub
            Else
                If strServerShare.EndsWith(IO.Path.DirectorySeparatorChar) Then
                    strServerShare = strServerShare.TrimEnd(IO.Path.DirectorySeparatorChar)
                End If
            End If

            ' check if user has rights to create directory / file

            ' if directory not present, create directory
            If IO.Directory.Exists(strDirPath) Then
                dirInfo = New IO.DirectoryInfo(strDirPath)
            Else
                dirInfo = IO.Directory.CreateDirectory(strDirPath)
            End If

            ' check if file/zip exists on local machine
            If Dir(strDirPath + IO.Path.DirectorySeparatorChar + strFileName).Length > 0 Then
                bolFileExists = True
            End If

            If Dir(strDirPath + IO.Path.DirectorySeparatorChar + strZipFileName).Length > 0 Then
                bolZipExists = True
            End If


            ' get file from server
            If Not bolZipExists And Not bolFileExists Then
                ' copy zip file from server to local machine
                Dim code(0) As String
                code(0) = "BeginSyncprocess"

                '  If Dir(strServerShare + IO.Path.DirectorySeparatorChar + strZipFileName).Length > 0 Then
                If Dir(strServerShare + IO.Path.DirectorySeparatorChar + strFileName).Length > 0 Then
                    ' delete db zip/file
                    RaiseEvent StartProgressScreen("Data Synchronization - Preparing Sync", 3, 0, "Getting Files from Server", code)
                    Dim delFiles = Dir(strDirPath + IO.Path.DirectorySeparatorChar + "*.BAK")
                    Do
                        If delFiles <> String.Empty Then
                            IO.File.Delete(strDirPath + IO.Path.DirectorySeparatorChar + delFiles)
                            delFiles = Dir()
                        End If

                    Loop Until delFiles = String.Empty

                    RaiseEvent FireProgressMessage("Getting Files from Server", 1, code(0))

                    delFiles = Dir(strDirPath + IO.Path.DirectorySeparatorChar + "*.ZIP")
                    Do
                        If delFiles <> String.Empty Then
                            IO.File.Delete(strDirPath + IO.Path.DirectorySeparatorChar + delFiles)
                            delFiles = Dir()
                        End If

                    Loop Until delFiles = String.Empty

                    RaiseEvent FireProgressMessage("Getting Files from Server", 2, code(0))


                    '       IO.File.Copy(strServerShare + IO.Path.DirectorySeparatorChar + Dir(strServerShare + IO.Path.DirectorySeparatorChar + strZipFileName), strDirPath + IO.Path.DirectorySeparatorChar + Dir(strServerShare + IO.Path.DirectorySeparatorChar + strZipFileName))
                    IO.File.Copy(strServerShare + IO.Path.DirectorySeparatorChar + Dir(strServerShare + IO.Path.DirectorySeparatorChar + strFileName), strDirPath + IO.Path.DirectorySeparatorChar + Dir(strServerShare + IO.Path.DirectorySeparatorChar + strFileName))
                    ' thiszipfilename = Dir(strServerShare + IO.Path.DirectorySeparatorChar + strZipFileName)
                    thiszipfilename = Dir(strServerShare + IO.Path.DirectorySeparatorChar + strFileName)
                    ' after getting the file from server set bool to true to extract the file(s) below
                    bolZipExists = True

                    RaiseEvent FireProgressMessage("Getting Files from Server", 3, code(0))

                    System.Threading.Thread.Sleep(1000)

                    RaiseEvent FireCloseProgressScreen()

                Else
                    MsgBox(String.Format("File not found. Please contact System Administrator {0}", String.Format("{0}{1}", vbCrLf, strServerShare + IO.Path.DirectorySeparatorChar + strFileName)))
                    Exit Sub
                End If
            End If

            ' extract zip
            '     If bolZipExists And Not bolFileExists AndAlso thiszipfilename.Length > 0 Then
            ' StartUnzipBackupFile(strDirPath, thiszipfilename)

            '   Else
            thiszipfilename = Dir(strDirPath + IO.Path.DirectorySeparatorChar + strFileName)
            PerformRestoreAfterUnZip(strDirPath + IO.Path.DirectorySeparatorChar + thiszipfilename)

            '   End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            dirInfo = Nothing
            LocalUserSettings.CurrentUser.SetValue("MusterSQLConnection", strConnStrPrev)
            Me.Cursor = Cursors.Default
        End Try
    End Sub

#End Region

    'Private Sub mnItemPhotosDocs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnItemPhotosDocs.Click
    '    Dim RegPhotoDocs As New DocsPhotos(Me)
    '    Try
    '        RegPhotoDocs.WindowState = FormWindowState.Maximized
    '        RegPhotoDocs.BringToFront()
    '        RegPhotoDocs.MdiParent = Me
    '        RegPhotoDocs.Show()
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub mnItemModuleColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnItemModuleColor.Click
        Dim x As MsgBoxResult
        Dim ColorName As String
        Dim ColorStruc As System.Drawing.Color

        x = Me.ColorPicker.ShowDialog()

        If x = MsgBoxResult.OK Then

            Try
                ColorName = ColorPicker.Color.Name
                If ColorPicker.Color.IsNamedColor Or ColorPicker.Color.IsKnownColor Then
                    ColorStruc = System.Drawing.Color.FromName(ColorName)
                    ColorName = ColorStruc.ToArgb().ToString
                End If
            Catch ex As Exception
                Dim errdisp As New ErrorReport(ex)
                errdisp.ShowDialog()
                Exit Sub
            End Try
            AppProfileInfo.Retrieve(AppUser.ID & "|MODULE_ID|" & strModuleTitle.ToUpper & "|HEADERCOLOR")
            AppProfileInfo.ProfileValue = ColorName
            AppProfileInfo.ModifiedBy = MusterContainer.AppUser.ID
            AppProfileInfo.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
        End If

        RefreshHeaderInfo()

    End Sub




#End Region

#Region "Form Events"


    Private Sub RegistrationServices_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Dim strMyName As String


        Try
            ' AddAdmin()
            If Not switchingUser Then
                MyGUID = New System.Guid


                Dim ThisForm As Logon
                ThisForm = New Logon(Me)
                ThisForm.ShowDialog()
                '
                ' Check if the user actually suceeded in logging in.
                '   If the appuser object doesn't exist, then user abandonded
                '
                txtOwnerQSKeyword.Focus()

                If AppUser Is Nothing Then
                    Me.Dispose()
                    Exit Sub
                End If
                If AppUser.ID = String.Empty Then
                    Me.Dispose()
                End If
            End If

            LoggedIn = True

            Try
                AppSemaphores = New MUSTER.BusinessLogic.pAppFlag(Me)
                AppProfileInfo = New MUSTER.BusinessLogic.pProfile
                ProfileData = New MUSTER.BusinessLogic.pProfile
                pCalendar = New MUSTER.BusinessLogic.pCalendar
                pFlag = New MUSTER.BusinessLogic.pFlag
                pOwn = New MUSTER.BusinessLogic.pOwner
                pEntity = New MUSTER.BusinessLogic.pEntity
                pLetter = New MUSTER.BusinessLogic.pLetter
                pConStruct = New MUSTER.BusinessLogic.pContactStruct
                pInspec = New MUSTER.BusinessLogic.pInspection
                oQS = New MUSTER.BusinessLogic.pSearch
                'Dim oPropType As New MUSTER.BusinessLogic.pPropertyType("Modules")
                'Dim dtTable As DataTable = oPropType.PropertiesTable
                'dtTable.DefaultView.RowFilter = "PropType_ID=" + oPropType.ID.ToString
                'cmbSearchModule.DataSource = dtTable.DefaultView  'oPropType.PropertiesTable
                'cmbSearchModule.ValueMember = "Property ID"
                'cmbSearchModule.DisplayMember = "Property Name"

                Dim dtTable As DataTable = AppUser.ListModulesUserCanSearch(AppUser.UserKey)
                dtTable.DefaultView.Sort = "PROPERTY_NAME"
                cmbSearchModule.DataSource = dtTable.DefaultView
                cmbSearchModule.ValueMember = "PROPERTY_ID"
                cmbSearchModule.DisplayMember = "PROPERTY_NAME"

                CheckMenuItemRights(AppUser.ListModulesUserHasAccessTo(AppUser.UserKey))

                bolLoadingForm = False


                If Not cmbSearchModule.DataSource Is Nothing Then
                    cmbSearchModule.SelectedIndex = -1
                    UIUtilsGen.SetComboboxItemByValue(cmbSearchModule, AppUser.DefaultModule)
                End If
                'cmbSearchModule.SelectedValue = Integer.Parse(AppUser.DefaultModule)
                If Not cmbQuickSearchFilter.DataSource Is Nothing Then
                    cmbQuickSearchFilter.SelectedIndex = 1
                End If

                ' CheckBox1.Checked = True

            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
                Me.Dispose()
            Finally
                bolLoadingForm = False
            End Try

            Try
                Dim MyConn As String = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")
                Dim MyConnEl() As String = MyConn.Split(";")
                Dim WorkTuple As String
                If AppSemaphores Is Nothing Then Exit Sub
                AppSemaphores.Retrieve("Connect", "String", LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection"))
                For Each WorkTuple In MyConnEl
                    Dim WorkEl = WorkTuple.Split("=")
                    If WorkEl(0) <> String.Empty Then
                        AppSemaphores.Retrieve(WorkEl(0), "", WorkEl(1))
                    End If
                Next
                'Dim oFilePath As New MUSTER.BusinessLogic.pFilePaths
                'oFilePath.Retrieve(UIUtilsGen.FilePathKey_Templates)
                'AppSemaphores.Retrieve("Template Location", "", oFilePath.FilePath)
                ''AppSemaphores.Retrieve("Report Location", "", "\\" + AppSemaphores.GetValuePair("Data Source", "") + "\" + strMusterShare + "\")
            Catch ex As Exception
                Throw ex
            End Try

            Me.txtUserWelcome.Text = "Welcome " + AppUser.ID
            AppUser.LogEntry("MAIN", MyGUID.ToString)

            ProfileData.GetAll()

            ''''''''''''''''''''''''''''''''''''''''
            'MONTH CALENDAR
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            Me.Update()
            LockWindowUpdate(CLng(Me.Handle.ToInt64))
            pCalendar.GetCalendarAll(AppUser.ID, False)

            LoadViewEntries()

            Try
                bolAllowDateChangedEvent = True
                bolCalLoad = True
                calTechnicalMonth.SetDate(Now())
            Finally
                bolAllowDateChangedEvent = False
                bolCalLoad = False
            End Try

            rdCalDay.Checked = True

            nGridWidth = dgCalToDo.Size.Width
            nGridHeight = dgDueToMe.Size.Height

            calendarRefreshTimer.Enabled = True
            calendarRefreshTimer.Stop()
            calendarRefreshTimer.Start()

            LockWindowUpdate(CLng(0))

            'Loading all the Documents into Letters Collection.
            'pLetter.GetAll(MusterContainer.AppUser.ID)

            CheckMouseToVisibleCalendar(True)
            Me.Text = "Muster - " + Me.AppSemaphores.GetValuePair("Data Source", "") + " - " + Me.AppSemaphores.GetValuePair("Initial Catalog", "")

            'Start Tickler 
            RaiseEvent StartTicklerScreen(Me.AppUser.ID)

            UpdateTicklerButton()

            TicklerTimerThread = New System.Timers.Timer(1000)

            TicklerTimerThread.Start()

            TicklerAutoOpenThread = New System.Timers.Timer(3000)

            TicklerAutoOpenThread.Start()


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            If Not AppUser Is Nothing Then
                Me.Show()
            End If
        End Try


    End Sub


    Sub OpenWindow(ByVal sender As Object, ByVal e As EventArgs)

        For Each child As Form In _container.MdiChildren
            If DirectCast(sender, MenuItem).Text = child.Text Then
                If Not child.Visible Then
                    child.Dock = DockStyle.Fill
                    child.Show()
                End If

                child.BringToFront()

            End If
        Next

    End Sub

    Sub MenuPopup(ByVal sender As Object, ByVal e As EventArgs) Handles mnItemWindows.Popup



        Dim keys As New Collections.ArrayList

        'finds old menu except close all windows
        While mnItemWindows.MenuItems.Count > 0
            mnItemWindows.MenuItems.RemoveAt(0)
        End While



        mnItemWindows.MenuItems.Add(Me.MenuItem1)
        mnItemWindows.MenuItems.Add(Me.MenuItem2)



        'adds physical windows
        Dim cnt As Integer = 0

        For Each child As Form In _container.MdiChildren
            If child.Visible Then
                mnItemWindows.MenuItems.Add(String.Format("{0}", child.Text), AddressOf OpenWindow)
                keys.Add(child.Text.Replace("  ", " ").Trim)
                cnt += 1
            End If

        Next

        'add additional windows (frames)
        For Each item2 As MenuItem In MnHistoryItems.MenuItems

            If Not keys.Contains(item2.Text.Replace("  ", " ").Trim) Then
                mnItemWindows.MenuItems.Add(item2.CloneMenu)
            End If
        Next


        MenuItem1.Index = MenuItem1.Parent.MenuItems.Count - 1
        MenuItem2.Index = MenuItem2.Parent.MenuItems.Count - 2

        keys.Clear()
        keys = Nothing

    End Sub


    Sub UpdateTicklerButton()

        Dim cnt As Integer = 0

        If Not TicklerScreen Is Nothing Then
            cnt = TicklerScreen.UnReadCount()
        End If

        If cnt > 0 Then
            btnTickler.Text = String.Format("There are {0} unread/unprocessed tickler messages.", cnt)
        Else
            btnTickler.Text = "There are no unread tickler messages."

        End If


    End Sub

    Public Sub PerformTicklerCheck(ByVal sender As Object, ByVal e As Timers.ElapsedEventArgs) Handles TicklerTimerThread.Elapsed

        TicklerTimerThread.Stop()

        RaiseEvent StartTicklerScreen(Me.AppUser.ID)

        UpdateTicklerButton()

        TicklerTimerThread.Interval = 900000

        TicklerTimerThread.Start()

    End Sub

    Public Sub PerformAutoOpenCheck(ByVal sender As Object, ByVal e As Timers.ElapsedEventArgs) Handles TicklerAutoOpenThread.Elapsed


        TicklerAutoOpenThread.Stop()

        If AutoOpen Then

            AutoOpen = False
            btnQuickOwnerSearch_Click(BootStrap._container, New EventArgs)

        End If

        TicklerAutoOpenThread.Start()

    End Sub

    ''' Attaches the Sort methods to any child that is open in order to have all sort memory in all grids
    Private Sub ActivatedToAttachChildToSortInstance(ByVal sender As Object, ByVal e As EventArgs)

        GridMaster.GlobalInstance.AttachFormToInstance(DirectCast(sender, Control))

    End Sub



    Private Sub frmClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim frm As Form

        Try
            If Me.MdiChildren.Length <= 1 Then
                btnCloseForm.Enabled = False
                btnCloseAll.Enabled = False
                mnItemCloseAll.Enabled = False
                lblOwnerInfo.Text = ""
                lblFacilityID.Text = ""
                lblFacilityInfo.Text = ""
                'txtOwnerQSKeyword.Text = String.Empty
            End If
            ClearBarometer()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Protected Overrides Sub OnMdiChildActivate(ByVal e As System.EventArgs)
        Dim frm As Form
        Try
            btnCloseForm.Enabled = True
            btnCloseAll.Enabled = True
            mnItemCloseAll.Enabled = True
            For Each frm In Me.MdiChildren

                With frm

                    RemoveHandler .Closed, AddressOf frmClosed
                    AddHandler .Closed, AddressOf frmClosed

                    RemoveHandler .Load, AddressOf ActivatedToAttachChildToSortInstance
                    AddHandler .Load, AddressOf ActivatedToAttachChildToSortInstance

                    RemoveHandler .Enter, AddressOf ActivatedToAttachChildToSortInstance
                    AddHandler .Enter, AddressOf ActivatedToAttachChildToSortInstance

                End With

            Next

            If bolLoadingForm Then Exit Sub

            Me.lblModuleName.Text = ""
            Me.lblOwner.Visible = False
            Me.lnklblPrevForm.Visible = False
            Me.lblOwnerInfo.Visible = False
            Me.lblFacility.Visible = False
            Me.lnkLblNextForm.Visible = False
            Me.lblFacilityInfo.Visible = False
            Me.lblFacilityID.Visible = False
            Me.lblModuleName.Visible = False
            Me.pnlCommonReferenceArea.Visible = False
            RefreshHeaderInfo()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "UI Events"

    Private Sub lnklblPrevForm_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnklblPrevForm.LinkClicked
        mnuItemAddOwner_Click(sender, e)
    End Sub
    Private Sub lnkLblNextForm_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextForm.LinkClicked
        mnItemAddFacility_Click(sender, e)
    End Sub
    Private Sub btnCloseForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseForm.Click
        Try
            Dim tempChild As Form
            tempChild = Me.ActiveMdiChild
            If Not tempChild Is Nothing Then
                tempChild.Close()
            End If
            If Me.MdiChildren.Length <= 0 Then
                btnCloseForm.Enabled = False
                btnCloseAll.Enabled = False
                mnItemCloseAll.Enabled = False
                ClearBarometer()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCloseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloseAll.Click

        mnItemCloseAll_Click(sender, e)



    End Sub
    Private Sub lblSearchCollapse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblSearchCollapse.Click

        If pnlQuickSearch.Visible Then
            pnlQuickSearch.Visible = False
            lblSearchCollapse.Text = "+"
            lblCollapseText.Text = "Expand Search"
        Else
            pnlQuickSearch.Visible = True
            lblSearchCollapse.Text = "-"
            lblCollapseText.Text = "Collapse Search"
        End If
    End Sub

    Private Sub txtOwnerQSKeyword_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOwnerQSKeyword.Leave
        strOwnerQSKeyWord = txtOwnerQSKeyword.Text
    End Sub

    Private Sub txtOwnerQSKeyword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOwnerQSKeyword.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnQuickOwnerSearch_Click(sender, e)
        End If
    End Sub

    Private Sub cmbSearchModule_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbSearchModule.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnQuickOwnerSearch_Click(sender, e)
        End If
    End Sub

    Private Sub cmbQuickSearchFilter_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbQuickSearchFilter.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnQuickOwnerSearch_Click(sender, e)
        End If
    End Sub

    Private Sub lblFacilityID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblFacilityID.TextChanged
        txtOwnerQSKeyword.Text = lblFacilityID.Text
        If lblFacilityID.Text.Trim <> String.Empty Then
            cmbQuickSearchFilter.SelectedItem = "Facility ID"
            cmbQuickSearchFilter.SelectedIndex = 1
            'ElseIf lblOwnerInfo.Text <> String.Empty Then
            '    Dim ch As Char
            '    For c As Integer = 0 To lblOwnerInfo.Text.ToCharArray.Length - 1
            '        If Not IsNumeric(lblOwnerInfo.Text.Chars(c)) Then
            '            txtOwnerQSKeyword.Text = lblOwnerInfo.Text.Substring(0, c)
            '            cmbQuickSearchFilter.SelectedItem = "Owner ID"
            '            cmbQuickSearchFilter.SelectedIndex = 0
            '            Exit For
            '        End If
            '    Next
        End If
    End Sub

#End Region

#Region "External Events"
    Private Shared Sub AppSemaphores_EntityChanged(ByVal FormGUID As String, ByVal FormKey As String, ByRef MyForm As Windows.Forms.Form) Handles AppSemaphores.EntityChanged
        Dim MyID As String
        Dim MyForm2 As MusterContainer
        Dim objID, objName As Object

        If Not MyForm Is Nothing Then
            MyForm2 = CType(MyForm, MusterContainer)
            Select Case FormKey
                Case "OwnerID"
                    Try
                        objName = AppSemaphores.GetValuePair(FormGUID.ToString, "OwnerName")
                        objID = AppSemaphores.GetValuePair(FormGUID.ToString, FormKey)
                        MyForm2.lblOwnerInfo.Text = objID & " : " & vbCrLf & IIf(objName Is Nothing, objName, objName.Replace("&", "&&"))
                    Catch e As Exception
                        MyForm2.lblOwnerInfo.Text = ""
                    End Try
                Case "OwnerName"
                    Try
                        objID = AppSemaphores.GetValuePair(FormGUID.ToString, "OwnerID")
                        objName = AppSemaphores.GetValuePair(FormGUID.ToString, FormKey)
                        MyForm2.lblOwnerInfo.Text = objID & " : " & vbCrLf & IIf(objName Is Nothing, objName, objName.Replace("&", "&&"))
                    Catch e As Exception
                        MyForm2.lblOwnerInfo.Text = ""
                    End Try
                Case "OwnerAddress"
                    Try
                        objName = AppSemaphores.GetValuePair(FormGUID.ToString, "OwnerName")
                        objID = AppSemaphores.GetValuePair(FormGUID.ToString, "OwnerID")
                        MyForm2.lblOwnerInfo.Text = objID & " : " & vbCrLf & IIf(objName Is Nothing, objName, objName.Replace("&", "&&"))
                    Catch e As Exception
                        MyForm2.lblOwnerInfo.Text = ""
                    End Try
                Case "FacilityID"
                    Try
                        objName = AppSemaphores.GetValuePair(FormGUID.ToString, "FacilityName")
                        objID = AppSemaphores.GetValuePair(FormGUID.ToString, "FacilityID")
                        MyForm2.lblFacilityID.Text = IIf(objID < 0, "", objID)
                        MyForm2.lblFacilityInfo.Text = IIf(objID < 0, "", IIf(objName Is Nothing, objName, objName.Replace("&", "&&")))
                    Catch e As Exception
                        MyForm2.lblFacilityID.Text = ""
                        MyForm2.lblFacilityInfo.Text = ""
                    End Try
                Case "FacilityName"
                    Try
                        objName = AppSemaphores.GetValuePair(FormGUID.ToString, FormKey)
                        objID = AppSemaphores.GetValuePair(FormGUID.ToString, "FacilityID")
                        MyForm2.lblFacilityID.Text = IIf(objID < 0, "", objID)
                        MyForm2.lblFacilityInfo.Text = IIf(objID < 0, "", IIf(objName Is Nothing, objName, objName.Replace("&", "&&")))
                    Catch e As Exception
                        MyForm2.lblFacilityID.Text = ""
                        MyForm2.lblFacilityInfo.Text = ""
                    End Try
                Case "FacilityAddress"
                    Try
                        objName = AppSemaphores.GetValuePair(FormGUID.ToString, "FacilityName")
                        objID = AppSemaphores.GetValuePair(FormGUID.ToString, "FacilityID")
                        MyForm2.lblFacilityID.Text = IIf(objID < 0, "", objID)
                        MyForm2.lblFacilityInfo.Text = IIf(objID < 0, "", IIf(objName Is Nothing, objName, objName.Replace("&", "&&")))
                    Catch e As Exception
                        MyForm2.lblFacilityID.Text = ""
                        MyForm2.lblFacilityInfo.Text = ""
                    End Try
                Case "CompanyID"
                    Try
                        objName = AppSemaphores.GetValuePair(FormGUID.ToString, "CompanyName")
                        objID = AppSemaphores.GetValuePair(FormGUID.ToString, "CompanyID")
                        MyForm2.lblOwnerInfo.Text = objID & " : " & vbCrLf & IIf(objName Is Nothing, objName, objName.Replace("&", "&&"))
                    Catch e As Exception
                        MyForm2.lblOwnerInfo.Text = ""
                    End Try
                Case "CompanyName"
                    Try
                        objName = AppSemaphores.GetValuePair(FormGUID.ToString, "CompanyName")
                        objID = AppSemaphores.GetValuePair(FormGUID.ToString, "CompanyID")
                        MyForm2.lblOwnerInfo.Text = objID & " : " & vbCrLf & IIf(objName Is Nothing, objName, objName.Replace("&", "&&"))
                    Catch e As Exception
                        MyForm2.lblOwnerInfo.Text = ""
                    End Try
            End Select
        End If
    End Sub
    'Private Shared Sub AppSemaphores_ActiveFormChanged(ByVal FormGUID As String, ByRef MyForm As Windows.Forms.Form) Handles AppSemaphores.ActiveFormChanged
    'End Sub
    Private Shared Sub SwitchToWindow(ByVal WindowTitle As String, ByRef WhichInstance As Windows.Forms.Form) Handles AppSemaphores.ActivateWindow
        Dim frm As Form
        Dim sFrm As MusterContainer

        sFrm = CType(WhichInstance, MusterContainer)
        For Each frm In WhichInstance.MdiChildren
            If frm.Text = WindowTitle Then
                frm.Activate()
                frm.Size = WhichInstance.ClientSize
                frm.BringToFront()
                Exit For
            End If
        Next
        sFrm.ActivateDocsandFlagsButtons(WindowTitle, WhichInstance)
    End Sub
    Private Shared Sub ActivateDocsandFlagsButtons(ByVal WindowTitle As String, ByRef WhichInstance As Windows.Forms.Form) Handles AppSemaphores.ActivateMusterControls
        Dim sFrm As MusterContainer

        sFrm = CType(WhichInstance, MusterContainer)
        If WindowTitle.StartsWith("Registration") Then
            sFrm.EnableDisableControls(WindowTitle)
        End If
    End Sub

    Public Sub FlagsChanged(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String, Optional ByVal eventID As Integer = -1, Optional ByVal eventType As Integer = -1) 'Handles objRegister.FlagsChanged
        Try
            LoadBarometer(entityID, entityType, [Module], ParentFormText, eventID, eventType)
            ' no need to update calendar if updating flags. if calendar entry was created, updating the calendar is handled at that point
            'DisplayOnDateSelectedOrChangedOrViewEntries()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub x_SearchResultSelection(ByVal nOwnerID As Integer, ByVal nFacilityId As Integer, ByVal Search_Type As String) Handles x.SearchResultSelection
        Try
            Dim localGUID As String
            Dim ChildForm As Windows.Forms.Form
            Dim TechForm As MUSTER.Technical
            Dim ClosrForm As MUSTER.Closure
            Dim FinancialForm As MUSTER.Financial
            Dim FeesForm As MUSTER.Fees
            Dim CandEForm As MUSTER.CandE
            Dim InspectionForm As MUSTER.Inspection

            Dim fGUID As System.Guid
            Dim BolRegForm As Boolean = False
            Dim BolTechForm As Boolean = False
            Dim bolClosrForm As Boolean = False
            Dim bolFinancialForm As Boolean = False
            Dim bolFeesForm As Boolean = False
            Dim bolInspection As Boolean = False

            Dim nEventID As Integer = 0
            Dim nClosureID As Integer = 0

            If Not frmAdvanceSearch Is Nothing Then
                If frmAdvanceSearch.strModule <> String.Empty Then
                    cmbSearchModule.Text = frmAdvanceSearch.strModule
                End If
            End If

            If Not Search_Type Is Nothing Then
                If cmbSearchModule.Text.Trim.ToUpper <> "INSPECTION" Then
                    If (Search_Type.IndexOf("Facility") > -1) Or (Search_Type.IndexOf("Ensite AIID") > -1) Then
                        If cmbSearchModule.Text = "C & E" Then
                            localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("FacilityID", nFacilityId.ToString, "C " + "&" + " E")
                            If localGUID = String.Empty Or localGUID Is Nothing Then
                                localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("OwnerID", nOwnerID.ToString, "C " + "&" + " E")
                            End If
                        Else
                            localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("FacilityID", nFacilityId.ToString, cmbSearchModule.Text)
                            If localGUID = String.Empty Or localGUID Is Nothing Then
                                localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("OwnerID", nOwnerID.ToString, cmbSearchModule.Text)
                            End If
                        End If
                    Else
                        If cmbSearchModule.Text = "C & E" Then
                            localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("OwnerID", nOwnerID.ToString, "C " + "&" + " E")
                        Else
                            localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("OwnerID", nOwnerID.ToString, cmbSearchModule.Text)
                        End If
                    End If
                End If
                Select Case cmbSearchModule.Text
                    Case "Registration"
                        If localGUID Is Nothing Then localGUID = String.Empty
                        If Not localGUID = String.Empty Then  'Existing Owner or Facility 
                            For Each ChildForm In Me.MdiChildren
                                If ChildForm.GetType.Name = "Registration" Then
                                    fGUID = CType(ChildForm, Registration).MyGuid
                                    If fGUID.ToString = localGUID Then
                                        ChildForm.Activate()
                                        objRegister = CType(ChildForm, Registration)
                                        If objRegister.lblOwnerIDValue.Text.Trim = nOwnerID.ToString Or (objRegister.lblFacilityIDValue.Text.Trim = nFacilityId.ToString) Then
                                            FillRegForm(nOwnerID, nFacilityId, Search_Type, objRegister)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            'New Registration 
                            If Search_Type.IndexOf("Owner") > -1 Then
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                'objRegister = New Registration(pOwn, "Owner", nOwnerID, 0)
                                objRegister = New Registration(pOwn, nOwnerID, 0)
                            Else
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                'pOwn.Facilities.Retrieve(pOwn.OwnerInfo, nFacilityId, , "FACILITY", , True)
                                'objRegister = New Registration(pOwn, "Facility", nOwnerID, nFacilityId)
                                objRegister = New Registration(pOwn, nOwnerID, nFacilityId)
                            End If
                            objRegister.MdiParent = Me
                            objRegister.Show()
                            objRegister.BringToFront()
                            ClearComboBox(objRegister)
                            objRegister.Tag = "Registration"
                        End If
                    Case "Technical"
                        If Not localGUID = String.Empty Then
                            'Existing Tech  Window 
                            For Each ChildForm In Me.MdiChildren
                                If ChildForm.GetType.Name = "Technical" Then
                                    fGUID = CType(ChildForm, Technical).MyGuid
                                    If fGUID.ToString = localGUID Then
                                        ChildForm.Activate()
                                        TechForm = CType(ChildForm, Technical)
                                        If TechForm.lblOwnerIDValue.Text.Trim = nOwnerID.ToString Or (TechForm.lblFacilityIDValue.Text.Trim = nFacilityId.ToString) Then
                                            FillRegForm(nOwnerID, nFacilityId, Search_Type, Nothing, TechForm)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            'New Tech Information.. 
                            If Search_Type.IndexOf("Owner") > -1 Then
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                TechForm = New Technical(nOwnerID, 0, pOwn)
                            Else
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                'pOwn.Facilities.Retrieve(pOwn.OwnerInfo, nFacilityId, , "FACILITY", , True)
                                If Not frmAdvanceSearch Is Nothing Then
                                    nEventID = frmAdvanceSearch.nEventID
                                End If
                                If nEventID > 0 Then
                                    TechForm = New Technical(nOwnerID, nFacilityId, pOwn, nEventID)
                                Else
                                    TechForm = New Technical(nOwnerID, nFacilityId, pOwn)
                                End If

                            End If
                            TechForm.MdiParent = Me
                            TechForm.Show()
                            TechForm.BringToFront()
                            ClearComboBox(TechForm)
                            TechForm.Tag = "Technical"
                        End If
                    Case "Closure"
                        If Not localGUID = String.Empty Then
                            'Existing Tech  Window 
                            For Each ChildForm In Me.MdiChildren
                                If ChildForm.GetType.Name = "Closure" Then
                                    fGUID = CType(ChildForm, Closure).MyGuid
                                    If fGUID.ToString = localGUID Then
                                        ChildForm.Activate()
                                        ClosrForm = CType(ChildForm, Closure)
                                        If ClosrForm.lblOwnerIDValue.Text.Trim = nOwnerID.ToString Or (ClosrForm.lblFacilityIDValue.Text.Trim = nFacilityId.ToString) Then
                                            FillRegForm(nOwnerID, nFacilityId, Search_Type, Nothing, Nothing, ClosrForm)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            'New Tech Information.. 
                            If Search_Type.IndexOf("Owner") > -1 Then
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                ClosrForm = New Closure(pOwn, nOwnerID, 0)
                            Else
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                If Not frmAdvanceSearch Is Nothing Then
                                    nClosureID = frmAdvanceSearch.nClosureID
                                End If
                                If nClosureID > 0 Then
                                    ClosrForm = New Closure(pOwn, nOwnerID, nFacilityId, nClosureID)
                                Else
                                    ClosrForm = New Closure(pOwn, nOwnerID, nFacilityId)
                                End If
                                If ClosrForm.dgClosureFacilityDetails.Rows.Count > 0 Then
                                    Me.mnuCancelAllOverDue.Enabled = True
                                Else
                                    Me.mnuCancelAllOverDue.Enabled = False
                                End If
                            End If
                            ClosrForm.MdiParent = Me
                            ClosrForm.Show()
                            ClosrForm.BringToFront()
                            ClearComboBox(ClosrForm)
                            ClosrForm.Tag = "Closure"
                        End If
                    Case "Financial"
                        If Not localGUID = String.Empty Then
                            'Existing Tech  Window 
                            For Each ChildForm In Me.MdiChildren
                                If ChildForm.GetType.Name = "Financial" Then
                                    fGUID = CType(ChildForm, Financial).MyGuid
                                    If fGUID.ToString = localGUID Then
                                        ChildForm.Activate()
                                        FinancialForm = CType(ChildForm, Financial)
                                        If FinancialForm.lblOwnerIDValue.Text.Trim = nOwnerID.ToString Or (FinancialForm.lblFacilityIDValue.Text.Trim = nFacilityId.ToString) Then
                                            FillRegForm(nOwnerID, nFacilityId, Search_Type, Nothing, Nothing, Nothing, FinancialForm)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            'New Tech Information.. 
                            If Search_Type.IndexOf("Owner") > -1 Then
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                FinancialForm = New Financial(pOwn, nOwnerID, 0)
                            Else
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                FinancialForm = New Financial(pOwn, nOwnerID, nFacilityId)

                            End If
                            FinancialForm.MdiParent = Me
                            FinancialForm.Show()
                            FinancialForm.BringToFront()
                            ClearComboBox(FinancialForm)
                            FinancialForm.Tag = "Financial"
                        End If
                    Case "Fees"
                        If Not localGUID = String.Empty Then
                            'Existing Tech  Window 
                            For Each ChildForm In Me.MdiChildren
                                If ChildForm.GetType.Name = "Fees" Then
                                    fGUID = CType(ChildForm, Fees).MyGuid
                                    If fGUID.ToString = localGUID Then
                                        ChildForm.Activate()
                                        FeesForm = CType(ChildForm, Fees)
                                        If FeesForm.lblOwnerIDValue.Text.Trim = nOwnerID.ToString Or (FeesForm.lblFacilityIDValue.Text.Trim = nFacilityId.ToString) Then
                                            FillRegForm(nOwnerID, nFacilityId, Search_Type, Nothing, Nothing, Nothing, Nothing, FeesForm)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            'New Tech Information.. 
                            If Search_Type.IndexOf("Owner") > -1 Then
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                FeesForm = New Fees(pOwn, nOwnerID, 0)
                            Else
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                FeesForm = New Fees(pOwn, nOwnerID, nFacilityId)
                            End If
                            FeesForm.MdiParent = Me
                            FeesForm.Show()
                            FeesForm.BringToFront()
                            ClearComboBox(FeesForm)
                            FeesForm.Tag = "Fees"
                        End If

                    Case "C & E"
                        If localGUID Is Nothing Then localGUID = String.Empty
                        If Not localGUID = String.Empty Then  'Existing Owner or Facility 
                            For Each ChildForm In Me.MdiChildren
                                If ChildForm.GetType.Name = "CandE" Then
                                    fGUID = CType(ChildForm, CandE).MyGuid
                                    If fGUID.ToString = localGUID Then
                                        ChildForm.Activate()
                                        CandEForm = CType(ChildForm, CandE)
                                        If CandEForm.lblOwnerIDValue.Text.Trim = nOwnerID.ToString Or (CandEForm.lblFacilityIDValue.Text.Trim = nFacilityId.ToString) Then
                                            FillRegForm(nOwnerID, nFacilityId, Search_Type, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, CandEForm)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            'New C & E
                            If Search_Type.IndexOf("Owner") > -1 Then
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                CandEForm = New CandE(pOwn, nOwnerID, 0)
                            Else
                                'If Not pOwn.OwnerCollection.Contains(nOwnerID) Then
                                '    pOwn.RetrieveAll(nOwnerID, cmbSearchModule.Text)
                                'Else
                                '    pOwn.Retrieve(nOwnerID, , , True)
                                'End If
                                'pOwn.Facilities.Retrieve(pOwn.OwnerInfo, nFacilityId, , "FACILITY", , True)
                                CandEForm = New CandE(pOwn, nOwnerID, nFacilityId)
                            End If
                            CandEForm.MdiParent = Me
                            CandEForm.Show()
                            CandEForm.BringToFront()
                            CandEForm.Tag = "C " + "&" + " E"
                        End If
                    Case "Inspection"
                        If localGUID Is Nothing Then localGUID = String.Empty
                        localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("WindowName", "Inspection Schedule", "Inspection")
                        If Not localGUID = String.Empty Then  'Existing inspection
                            'Existing Inspection Window
                            For Each ChildForm In Me.MdiChildren
                                If ChildForm.GetType.Name = "Inspection" Then
                                    fGUID = CType(ChildForm, Inspection).MyGuid
                                    If fGUID.ToString = localGUID Then
                                        ChildForm.Activate()
                                        InspectionForm = CType(ChildForm, Inspection)
                                        FillRegForm(nOwnerID, nFacilityId, Search_Type, , , , , , InspectionForm)
                                        Exit Sub
                                    End If
                                End If
                            Next
                        Else
                            'New Inspection
                            InspectionForm = New Inspection(pInspec, nOwnerID, nFacilityId)
                            InspectionForm.MdiParent = Me
                            InspectionForm.Show()
                            InspectionForm.BringToFront()
                            InspectionForm.Tag = "Inspection"
                            HideShowOwnerInfo(False)
                        End If
                End Select
            Else
                MessageBox.Show("No Records Found.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub x_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles x.Closed
        x = Nothing
    End Sub
    Private Sub x_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles x.Closing
        x = Nothing
    End Sub
    Private Sub SearchResultsCount(ByVal nCount As Integer, ByVal strSrc As String) Handles x.SearchResults
        If nCount > 1 Then
            x.BringToFront()
            x.MdiParent = Me
            x.Show()
        End If
    End Sub
    Private Sub CompanySearchResultsCount(ByVal nCount As Integer, ByVal strSrc As String) Handles cComSearch.CompanySearchResults
        If nCount > 1 Then
            cComSearch.BringToFront()
            cComSearch.MdiParent = Me
            cComSearch.Show()
        End If
    End Sub
    Private Sub Company_SearchResultSelection(ByVal nCompanyID As Integer, ByVal nLicenseeID As Integer, ByVal nCompAddressID As Integer, ByVal Search_Type As String) Handles cComSearch.CompanySearchSelection

        Try
            Dim localGUID As String
            Dim ChildForm As Windows.Forms.Form
            Dim ComForm As MUSTER.Company
            Dim LicForm As MUSTER.Licensees
            Dim MgrForm As MUSTER.Managers
            Dim fGUID As System.Guid

            If Not Search_Type Is Nothing Then

                If Search_Type.IndexOf("Company Name") > -1 Or nCompanyID <> 0 Then
                    localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("CompanyID", nCompanyID.ToString, cmbSearchModule.Text)
                Else
                    localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("LicenseeID", nLicenseeID.ToString, cmbSearchModule.Text)
                End If

                Select Case Search_Type
                    Case "Company Name"
                        If localGUID Is Nothing Then localGUID = String.Empty
                        If Not localGUID = String.Empty Then  'Existing Company 
                            For Each ChildForm In Me.MdiChildren
                                If ChildForm.GetType.Name = "Company" Then
                                    fGUID = CType(ChildForm, Company).MyGuid
                                    If fGUID.ToString = localGUID Then
                                        ChildForm.Activate()
                                        ComForm = CType(ChildForm, Company)
                                        If ComForm.txtCompanyID.Text.Trim = nCompanyID.ToString Then
                                            ComForm.MdiParent = Me
                                            ComForm.Show()
                                            ComForm.Tag = "Company"
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            'New Company
                            ComForm = New Company(nCompanyID)
                            ComForm.MdiParent = Me
                            ComForm.Show()
                            ComForm.BringToFront()
                            ComForm.Tag = "Company"
                        End If
                    Case "Licensee Name"
                        If nCompanyID <> 0 Then
                            If localGUID Is Nothing Then localGUID = String.Empty
                            If Not localGUID = String.Empty Then  'Existing Company 
                                For Each ChildForm In Me.MdiChildren
                                    If ChildForm.GetType.Name = "Company" Then
                                        fGUID = CType(ChildForm, Company).MyGuid
                                        If fGUID.ToString = localGUID Then
                                            ChildForm.Activate()
                                            ComForm = CType(ChildForm, Company)
                                            If ComForm.txtCompanyID.Text.Trim = nCompanyID.ToString Then
                                                ComForm.MdiParent = Me
                                                ComForm.Show()
                                                ComForm.BringToFront()
                                                ComForm.Tag = "Company"
                                                ComForm.Licensees(nLicenseeID)
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                'New Company
                                ComForm = New Company(nCompanyID)
                                ComForm.MdiParent = Me
                                ComForm.Show()
                                ComForm.BringToFront()
                                ComForm.Tag = "Company"
                                ComForm.Licensees(nLicenseeID)
                            End If
                        Else
                            If localGUID Is Nothing Then localGUID = String.Empty
                            If Not localGUID = String.Empty Then  'Existing Licensee
                                For Each ChildForm In Me.MdiChildren
                                    If ChildForm.GetType.Name = "Licensees" Then
                                        fGUID = CType(ChildForm, Licensees).MyGuid
                                        If fGUID.ToString = localGUID Then
                                            ChildForm.Activate()
                                            LicForm = CType(ChildForm, Licensees)
                                            If LicForm.LicenseeID = nLicenseeID Then
                                                LicForm.MdiParent = Me
                                                LicForm.Show()
                                                LicForm.Tag = "Licensee"
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                'New Licensee 
                                LicForm = New Licensees(nLicenseeID, nCompanyID, nCompAddressID, )
                                LicForm.MdiParent = Me
                                LicForm.Show()
                                LicForm.Tag = "Licensees"
                            End If
                        End If
                    Case "Manager Name"
                        If nCompanyID <> 0 Then
                            If localGUID Is Nothing Then localGUID = String.Empty
                            If Not localGUID = String.Empty Then  'Existing Company 
                                For Each ChildForm In Me.MdiChildren
                                    If ChildForm.GetType.Name = "Company" Then
                                        fGUID = CType(ChildForm, Company).MyGuid
                                        If fGUID.ToString = localGUID Then
                                            ChildForm.Activate()
                                            ComForm = CType(ChildForm, Company)
                                            If ComForm.txtCompanyID.Text.Trim = nCompanyID.ToString Then
                                                ComForm.MdiParent = Me
                                                ComForm.Show()
                                                ComForm.BringToFront()
                                                ComForm.Tag = "Company"
                                                ComForm.Managers(nLicenseeID)
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                'New Company
                                ComForm = New Company(nCompanyID)
                                ComForm.MdiParent = Me
                                ComForm.Show()
                                ComForm.BringToFront()
                                ComForm.Tag = "Company"
                                ComForm.Managers(nLicenseeID)
                            End If
                        Else
                            If localGUID Is Nothing Then localGUID = String.Empty
                            If Not localGUID = String.Empty Then  'Existing Manager
                                For Each ChildForm In Me.MdiChildren
                                    If ChildForm.GetType.Name = "Managers" Then
                                        fGUID = CType(ChildForm, Managers).MyGuid
                                        If fGUID.ToString = localGUID Then
                                            ChildForm.Activate()
                                            MgrForm = CType(ChildForm, Managers)
                                            If MgrForm.ManagerID = nLicenseeID Then
                                                MgrForm.MdiParent = Me
                                                MgrForm.Show()
                                                MgrForm.Tag = "Manager"
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                'New Manager 
                                MgrForm = New Managers(nLicenseeID, nCompanyID, nCompAddressID, )
                                MgrForm.MdiParent = Me
                                MgrForm.Show()
                                MgrForm.Tag = "Managers"
                            End If
                        End If
                End Select
            Else
                MessageBox.Show("No Records Found.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Private Sub cComSearch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cComSearch.Closing
        cComSearch = Nothing
    End Sub
    Private Sub cComSearch_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles cComSearch.Closed
        cComSearch = Nothing
    End Sub
#End Region

#Region "Calendar"

#Region "Button Events"
    Private Sub btnCalNewEntry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalNewEntry.Click
        Dim strGUID As String

        Me.AppSemaphores.Retrieve("New Calender Entry")
        If Not Me.AppSemaphores.Value Is Nothing Then
            strGUID = Me.AppSemaphores.Value.ToString
        Else
            strGUID = String.Empty
        End If

        If strGUID <> String.Empty Then
            Me.AppSemaphores.MakeWindowActive("New Calendar Entry")
            Exit Sub
        End If
        Dim frmCal As New Calendar(IIf(cmbViewCalEntries.Text = String.Empty, AppUser.ID, cmbViewCalEntries.Text), Me)
        Try
            AddHandler frmCal.Closing, AddressOf frmCalendarClosing
            AddHandler frmCal.Closed, AddressOf frmCalendarClosed
            frmCal.bolLoading = True
            frmCal.strMode = "ADD"
            frmCal.BringToFront()
            frmCal.WindowState = FormWindowState.Maximized
            frmCal.Show()
            'frmCal.cmbCalTargetGroup.DataSource = MusterContainer.AppUser.ListAllGroups
            'frmCal.cmbCalTargetGroup.DisplayMember = "USER_GROUP"
            'frmCal.cmbCalTargetGroup.ValueMember = "USER_GROUP"
            'frmCal.cmbCalTargetGroup.SelectedIndex = -1
            'frmCal.cmbCalTargetGroup.SelectedIndex = -1
            frmCal.bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCalToDoModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalToDoModify.Click
        Try
            ModifyCalendar(dgCalToDo)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub btnCalDueToMeModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalDueToMeModify.Click
        Try
            ModifyCalendar(dgDueToMe)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnToDoDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToDoDelete.Click
        Try
            DeleteCalendar(dgCalToDo)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
        End Try
    End Sub
    Private Sub btnDueToMeDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDueToMeDelete.Click
        Try

            DeleteCalendar(dgDueToMe)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
        End Try
    End Sub
    Private Sub btnToDoMarkCompleted_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToDoMarkCompleted.Click
        Try
            MarkCompleted(dgCalToDo)
            chkToDoCompItems_CheckedChanged(sender, e)
            chkDueToMeCompItems_CheckedChanged(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
        End Try
    End Sub
    Private Sub btnDueToMeMarkComp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDueToMeMarkComp.Click
        Try
            MarkCompleted(dgDueToMe)
            chkToDoCompItems_CheckedChanged(sender, e)
            chkDueToMeCompItems_CheckedChanged(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
        End Try
    End Sub




    Private Sub btnDock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDock.Click
        Try
            Me.pnlRightMost.Size = New System.Drawing.Size(16, 701)
            lockChildren()
            bolpnlVisibilityChanged = True
            If btnDock.Text = ">" Then
                btnDock.Text = "<"
                pnlRightContainer.Visible = False
                RightPanelDockFlag = False
            Else
                RightPanelDockFlag = True
                btnDock.Text = ">"
                pnlRightContainer.Visible = True
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            releaseChildren()
        End Try
    End Sub
#End Region

#Region "Common Procs and Functions"
    Private Sub BoldCalendarDates(ByVal dtCalendar As DataTable)
        Dim calDataRow As DataRow
        Dim dtEvent As DateTime
        Dim i As Integer = 0

        Try
            Dim count As Integer
            count = dtCalendar.Rows.Count

            Dim dtBoldDates(count) As System.DateTime

            For Each calDataRow In dtCalendar.Rows

                If Not calDataRow("DATE_DUE") Is Nothing Then
                    dtEvent = calDataRow("DATE_DUE").ToString
                    Dim dtDrMonth As Integer = dtEvent.Month
                    Dim dtDrYear As Integer = dtEvent.Year
                    Dim dtDrDay As Integer = dtEvent.Day
                    dtBoldDates(i) = New System.DateTime(dtDrYear, dtDrMonth, dtDrDay, 0, 0, 0, 0)
                    i = i + 1
                End If
            Next

            calTechnicalMonth.BoldedDates = dtBoldDates

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub BoldCalendarDates(ByVal dvCalendar As DataView)
        Dim calDataRow As DataRow
        Dim dtEvent As DateTime
        Dim i As Integer = 0

        Try
            Dim count As Integer
            count = dvCalendar.Count

            Dim dtBoldDates(count) As System.DateTime

            For j As Integer = 0 To dvCalendar.Count - 1
                calDataRow = dvCalendar.Item(j).Row
                If Not calDataRow("DATE_DUE") Is Nothing Then
                    dtEvent = calDataRow("DATE_DUE").ToString
                    Dim dtDrMonth As Integer = dtEvent.Month
                    Dim dtDrYear As Integer = dtEvent.Year
                    Dim dtDrDay As Integer = dtEvent.Day
                    dtBoldDates(i) = New System.DateTime(dtDrYear, dtDrMonth, dtDrDay, 0, 0, 0, 0)
                    i = i + 1
                End If
            Next

            calTechnicalMonth.BoldedDates = dtBoldDates

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub BoldDatesInCalendar(Optional ByVal strUser As String = "")
        Dim calDataRow As DataRow
        Dim dtEvent As DateTime
        'Dim calTask As New InfoRepository.CalendarTask
        Dim dtCal As DataTable
        Dim i As Integer = 0

        Try
            Dim count As Integer
            If strUser = String.Empty Then strUser = AppUser.ID

            dtCal = pCalendar.GetLoadedCalendar(strUser, "USER")
            count = dtCal.Rows.Count

            Dim dtBoldDates(count) As System.DateTime

            For Each calDataRow In dtCal.Rows

                If Not calDataRow("DATE_DUE") Is Nothing Then
                    dtEvent = calDataRow("DATE_DUE").ToString
                    Dim dtDrMonth As Integer = dtEvent.Month
                    Dim dtDrYear As Integer = dtEvent.Year
                    Dim dtDrDay As Integer = dtEvent.Day
                    dtBoldDates(i) = New System.DateTime(dtDrYear, dtDrMonth, dtDrDay, 0, 0, 0, 0)
                    i = i + 1
                End If
            Next

            calTechnicalMonth.BoldedDates = dtBoldDates

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub setDataGridWidth(ByVal dg As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nIndex As Integer)
        Dim ts As New DataGridTableStyle
        Dim nTaskWidth As Integer
        Dim nWidth As Integer

        Try

            'dg.DisplayLayout.Bands(0).Columns("DATE_DUE").PerformAutoResize()
            'dg.DisplayLayout.Bands(0).Columns("TARGET").PerformAutoResize()
            'dg.DisplayLayout.Bands(0).Columns("COMPLETED").PerformAutoResize()

            dg.DisplayLayout.Bands(0).Columns("DATE_DUE").Width = 65
            dg.DisplayLayout.Bands(0).Columns("TARGET").Width = 60
            dg.DisplayLayout.Bands(0).Columns("COMPLETED").Width = 40

            nWidth = calTechnicalMonth.Size.Width
            If nWidth > 192 Then
                nTaskWidth = (dg.Size.Width - (dg.DisplayLayout.Bands(0).Columns(0).Width) - (dg.DisplayLayout.Bands(0).Columns(1).Width) - (dg.DisplayLayout.Bands(0).Columns(3).Width))
                dg.DisplayLayout.Bands(0).Columns("TASK_DESCRIPTION").Width = nTaskWidth
            Else
                dg.DisplayLayout.Bands(0).Columns("TASK_DESCRIPTION").Width = 185
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub SetGridsReadOnly(ByVal nIndex As Integer, ByVal dsSet As DataSet)
        Dim dsCol As DataColumn
        Try
            For Each dsCol In dsSet.Tables(nIndex).Columns
                dsCol.ReadOnly = True
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub DisplayOnDateSelectedOrChangedOrViewEntries(Optional ByVal strUserId As String = Nothing, Optional ByVal strGroupId As String = Nothing)
        Dim sender As System.Object
        Dim e As System.EventArgs
        Try
            pCalendar.GetCalendarAll(AppUser.ID, False)
            chkToDoCompItems_CheckedChanged(sender, e)
            chkDueToMeCompItems_CheckedChanged(sender, e)
        Catch ex As Exception
            Throw ex
        Finally

        End Try

    End Sub
    Private Sub ColorInOverDueEntries(ByRef MyGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try
            Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim currentDate As Date = Today
            Dim strUser As String = AppUser.ID
            Dim dtAssignedGroups As DataTable
            Dim drGroup As DataRow
            Dim dueDate As Date
            If Not cmbViewCalEntries.SelectedIndex < 0 Then
                strUser = cmbViewCalEntries.Text.ToString
            End If
            For Each ugrow In MyGrid.Rows
                If Not ugrow.Cells("COMPLETED").Value = "True" Then

                    If ugrow.Cells("USER_ID").Value = strUser Then
                        dueDate = DateAdd(DateInterval.Day, 30, ugrow.Cells("DATE_DUE").Value)
                        If ugrow.Cells("SOURCE_USER_ID").Value <> ugrow.Cells("USER_ID").Value Then
                            ugrow.Appearance.BackColor = System.Drawing.Color.Aqua
                        End If
                        If currentDate.Date > dueDate.Date Then
                            ugrow.Appearance.BackColor = System.Drawing.Color.Red
                        End If
                    Else
                        If Not ugrow.Cells("GROUP_ID").Value = String.Empty Then
                            dtAssignedGroups = MusterContainer.AppUser.ListMemberships
                            For Each drGroup In dtAssignedGroups.Rows
                                If ugrow.Cells("GROUP_ID").Value = drGroup("USER_GROUP") Then
                                    dueDate = DateAdd(DateInterval.Day, 30, ugrow.Cells("DATE_DUE").Value)
                                    If currentDate.Date > dueDate.Date Then
                                        ugrow.Appearance.BackColor = System.Drawing.Color.Red
                                    End If
                                End If
                            Next
                        Else
                            ugrow.Appearance.BackColor = System.Drawing.Color.Gainsboro
                        End If
                    End If

                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ModifyCalendar(ByRef myGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim frmCal As New Calendar(IIf(cmbViewCalEntries.Text = String.Empty, AppUser.ID, cmbViewCalEntries.Text), Me)
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try

            If Not myGrid.Rows.Count > 0 Then Exit Sub

            If myGrid.ActiveRow Is Nothing Then
                MsgBox("Select the row to modify")
                Exit Sub
            Else

                If myGrid.ActiveRow.Cells("SOURCE_USER_ID").Text.ToString = "SYSTEM" Then
                    MsgBox("You Cannot Modify System Generated Calendar Entry.")
                    Exit Sub
                End If

                ugrow = myGrid.ActiveRow
                frmCal.bolLoading = True
                frmCal.dtPickCalendarDueDateValue.Format = DateTimePickerFormat.Custom
                frmCal.dtPickCalendarDueDateValue.CustomFormat = ugrow.Cells("DATE_DUE").Text
                frmCal.dtPickCalendarDueDateValue.Checked = True

                If ugrow.Cells("TO_DO").Text = True Then
                    frmCal.RdCalCategoryToDo.Checked = True
                    frmCal.rdCalCategoryDueToMe.Checked = False
                Else
                    frmCal.RdCalCategoryToDo.Checked = False
                    frmCal.rdCalCategoryDueToMe.Checked = True
                End If

                'frmCal.cmbCalTargetGroup.DataSource = MusterContainer.AppUser.ListAllGroups
                'frmCal.cmbCalTargetGroup.DisplayMember = "USER_GROUP"
                'frmCal.cmbCalTargetGroup.ValueMember = "USER_GROUP"
                'frmCal.cmbCalTargetGroup.SelectedIndex = -1

                If ugrow.Cells("USER_ID").Text <> "" Then
                    frmCal.rdCalTargetUser.Checked = True
                    frmCal.rdCalTargetGroup.Checked = False
                    frmCal.cmbCalTargetGroup.SelectedIndex = -1
                    If frmCal.cmbCalTargetGroup.SelectedIndex <> -1 Then
                        frmCal.cmbCalTargetGroup.SelectedIndex = -1
                    End If
                    frmCal.lblUser.Text = ugrow.Cells("USER_ID").Text
                Else
                    frmCal.rdCalTargetUser.Checked = False
                    frmCal.rdCalTargetGroup.Checked = True
                    frmCal.cmbCalTargetGroup.Text = ugrow.Cells("GROUP_ID").Text
                End If

                frmCal.txtCalDescription.Text = ugrow.Cells("TASK_DESCRIPTION").Text
                frmCal.txtCalDescription.Tag = ugrow.Cells("CALENDAR_INFO_ID").Text

                AddHandler frmCal.Closing, AddressOf frmCalendarClosing
                AddHandler frmCal.Closed, AddressOf frmCalendarClosed

                frmCal.strMode = "MODIFY"
                frmCal.BringToFront()
                frmCal.WindowState = FormWindowState.Maximized
                frmCal.Show()
                frmCal.bolLoading = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub DeleteCalendar(ByRef myGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim msgResult As MsgBoxResult
        Dim success As Boolean
        Try

            If Not myGrid.Rows.Count > 0 Then Exit Sub

            If myGrid.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete")
                Exit Sub
            Else
                If myGrid.ActiveRow.Cells("SOURCE_USER_ID").Text.ToString = "SYSTEM" Then
                    MsgBox("You Cannot Delete System Generated Calendar Entry.")
                    Exit Sub
                End If

                msgResult = MsgBox("Do you want to Delete this row.?", MsgBoxStyle.YesNo, "Calendar")
                If msgResult = MsgBoxResult.No Then
                    Exit Sub
                Else
                    'Retrieve based on ID.
                    pCalendar.Retrieve(Integer.Parse(myGrid.ActiveRow.Cells("CALENDAR_INFO_ID").Text))
                    pCalendar.Deleted = True
                    If pCalendar.CalendarId <= 0 Then
                        pCalendar.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pCalendar.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    'Mark Deleted in Database.
                    pCalendar.Save(CType(UIUtilsGen.ModuleID.[Global], Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    'Remove the Row from Collection.
                    pCalendar.Remove(Integer.Parse(myGrid.ActiveRow.Cells("CALENDAR_INFO_ID").Text))
                    ' To Remove the associated Flag Entry 
                    DeleteAssociatedFlags(myGrid)
                    myGrid.ActiveRow.Delete(False)
                End If
            End If

            MsgBox("The Record is Deleted Successfully.")

        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
    Private Sub MarkCompleted(ByRef myGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim msgResult As MsgBoxResult
        Dim success As Boolean
        Try
            If Not myGrid.Rows.Count > 0 Then Exit Sub

            If myGrid.ActiveRow Is Nothing Then
                MsgBox("Select row to Mark Completed")
                Exit Sub
            Else
                If myGrid.ActiveRow.Cells("SOURCE_USER_ID").Text.ToString = "SYSTEM" Then
                    If Not AppUser.HEAD_ADMIN Then
                        MsgBox("You Cannot Mark Complete System Generated Calendar Entry.")
                        Exit Sub
                    End If
                End If

                msgResult = MsgBox("Do you want to Mark this row Completed?", MsgBoxStyle.YesNo, "Calendar")
                If msgResult = MsgBoxResult.No Then
                    Exit Sub
                Else
                    pCalendar.Retrieve(Integer.Parse(myGrid.ActiveRow.Cells("CALENDAR_INFO_ID").Text))
                    pCalendar.Completed = True
                    If pCalendar.CalendarId <= 0 Then
                        pCalendar.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pCalendar.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    pCalendar.Save(CType(UIUtilsGen.ModuleID.[Global], Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    ' To Remove the associated Flag Entry 
                    DeleteAssociatedFlags(myGrid)

                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
    Private Sub DeleteAssociatedFlags(ByRef myGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try
            ' To Remove the associated Flag Entry 
            'MusterContainer.pFlag.FindFromCalendarID(Integer.Parse(myGrid.ActiveRow.Cells("CALENDAR_INFO_ID").Text))
            MusterContainer.pFlag.RetrieveFlags(, , , , , Integer.Parse(myGrid.ActiveRow.Cells("CALENDAR_INFO_ID").Text))

            For Each flagInfo As MUSTER.Info.FlagInfo In MusterContainer.pFlag.FlagsCol.Values
                flagInfo.Deleted = True
                flagInfo.ModifiedBy = MusterContainer.AppUser.ID
            Next
            MusterContainer.pFlag.Flush()
            MusterContainer.pFlag.FlagsCol = New MUSTER.Info.FlagsCollection
            LoadBarometer()

        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
    Private Sub HideGridColumns(ByRef myGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        myGrid.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        myGrid.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        myGrid.DisplayLayout.Bands(0).Columns("CALENDAR_INFO_ID").Width = 50
        myGrid.DisplayLayout.Bands(0).Columns("CALENDAR_INFO_ID").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("NOTIFICATION_DATE").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("CURRENT_COLOR_CODE").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("SOURCE_USER_ID").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("TO_DO").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("DUE_TO_ME").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("USER_ID").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("GROUP_ID").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("CREATED_BY").Hidden = True
        'myGrid.DisplayLayout.Bands(0).Columns("DATE_CREATED").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("LAST_EDITED_BY").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("Owning_Entity_Type").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("Owning_Entity_ID").Hidden = True
        myGrid.DisplayLayout.Bands(0).Columns("DATE_DUE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
    End Sub
    Friend Sub setFilter()
        Try
            Select Case rOption
                Case "M"
                    completedTaskFilterString = " DATE_DUE >= '" + filterStartMonth + " ' AND DATE_DUE <= '" + filterEndMonth + "'"
                    filterString = " DATE_DUE <= '" + filterEndMonth + "'"
                Case "W"
                    completedTaskFilterString = " DATE_DUE >= '" + filterStartWeek + " ' AND DATE_DUE <= '" + filterEndWeek + "'"
                    filterString = " DATE_DUE <= '" + filterEndWeek + "'"
                Case "D"
                    completedTaskFilterString = " DATE_DUE >= '" + filterStartMonth + " ' AND DATE_DUE <= '" + filterEndDay + "'"
                    filterString = "DATE_DUE <= '" + filterEndDay + "'"
            End Select

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadViewEntries()
        Dim dtSuperUsers As DataTable
        Dim drow As DataRow
        bolCalLoad = True

        ' To Load the Supervised Users for the Target User.
        ' #2358
        If AppUser.HEAD_ADMIN Then
            dtSuperUsers = AppUser.ListAllUsers(True, False)
        Else
            dtSuperUsers = AppUser.ListSupervisedUsers()
        End If
        If dtSuperUsers.Rows.Count > 0 Then
            cmbViewCalEntries.Enabled = True
            'Add Target user in the List.
            If Not AppUser.HEAD_ADMIN Then
                drow = dtSuperUsers.NewRow()
                drow("USER_ID") = AppUser.ID
                drow("USERNAME") = AppUser.Name
                dtSuperUsers.Rows.Add(drow)
            End If
            cmbViewCalEntries.DataSource = dtSuperUsers
            cmbViewCalEntries.DisplayMember = "USER_ID"
            cmbViewCalEntries.ValueMember = "USER_ID"
            cmbViewCalEntries.SelectedIndex = -1
            cmbViewCalEntries.SelectedIndex = -1
        Else
            cmbViewCalEntries.Enabled = False
        End If
        bolCalLoad = False
    End Sub
    Public Sub LoadCalendarforSelectedUser()
        Try
            dgCalToDo.DataSource = Nothing
            dgDueToMe.DataSource = Nothing

            'Reset both grid background colors
            dgCalToDo.DisplayLayout.Appearance.BackColor = Color.White
            dgDueToMe.DisplayLayout.Appearance.BackColor = Color.White

            Dim dtCal As DataTable = pCalendar.GetLoadedCalendar(cmbViewCalEntries.Text.ToString, "USER")

            ' set up the CalDueToMe
            'Dim dtDueToMe As DataTable
            'dtDueToMe = pCalendar.GetLoadedCalendar(cmbViewCalEntries.Text.ToString, "USER")
            Dim dvDueToMe As DataView = dtCal.DefaultView
            dvDueToMe.RowFilter = filterString + IIf(Not filterString Is Nothing, " AND COMPLETED=FALSE AND DUE_TO_ME = TRUE", "")
            dvDueToMe.Sort = "DATE_DUE ASC"
            dgDueToMe.DataSource = dvDueToMe
            BoldCalendarDates(dvDueToMe)
            HideGridColumns(dgDueToMe)
            setDataGridWidth(dgDueToMe, 0)
            ColorInOverDueEntries(dgDueToMe)

            ' set up the CalToDo
            'Dim dtToDo As DataTable
            'dtToDo = pCalendar.GetLoadedCalendar(cmbViewCalEntries.Text.ToString, "USER")
            Dim dvToDo As DataView = dtCal.DefaultView
            dvToDo.RowFilter = filterString + IIf(Not filterString Is Nothing, " AND COMPLETED=FALSE AND TO_DO = TRUE", "")
            dvToDo.Sort = "DATE_DUE ASC"
            dgCalToDo.DataSource = dvToDo
            BoldCalendarDates(dvToDo)
            HideGridColumns(dgCalToDo)
            setDataGridWidth(dgCalToDo, 1)
            ColorInOverDueEntries(dgCalToDo)
        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
    Private Function ValidateUser(ByRef ugrid As Infragistics.Win.UltraWinGrid.UltraGrid) As Boolean
        Dim dtSuperUsers As DataTable
        Dim drow As DataRow

        If ugrid.Rows.Count <= 0 Then Return False

        If ugrid.ActiveRow Is Nothing Then Return False

        If ugrid.ActiveRow.Cells("SOURCE_USER_ID").Value = "SYSTEM" Then
            'Return False
            ' #2358
            If AppUser.HEAD_ADMIN OrElse AppUser.ID = "Donna" Then
                Return True
            Else
                Return False
            End If
            Exit Function
        End If

        If AppUser.ID = ugrid.ActiveRow.Cells("TARGET").Value Then
            Return True
            Exit Function
        Else

            ' To Check the selected user is the Supervisor for Login User.
            dtSuperUsers = AppUser.ListSupervisedUsers()
            If dtSuperUsers.Rows.Count > 0 Then
                For Each drow In dtSuperUsers.Rows
                    If drow("USER_ID") = ugrid.ActiveRow.Cells("TARGET").Value Then
                        Return True
                        Exit For
                    End If
                Next
            End If

            ' To Check the selected user is in the Login User's group.
            dtSuperUsers = AppUser.ListMemberships()
            If dtSuperUsers.Rows.Count > 0 Then
                Return True
            End If

        End If
        Return False
    End Function
    Private Sub lockChildren()
        Dim frm As Form
        Try
            For Each frm In Me.MdiChildren
                frm.Update()
                If frm.Visible Then LockWindowUpdate(frm.Handle.ToInt64)
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub releaseChildren()
        Dim frm As Form
        Try
            For Each frm In Me.MdiChildren
                LockWindowUpdate(CLng(0))
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Form and Other Events"
    Private Sub tbCtrlRightPane_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlRightPane.Resize
        Try
            pnlCalToDoGrid.Width = tbCtrlRightPane.Width
            dgCalToDo.Width = pnlCalToDoGrid.Width
            pnlDueToMeGrid.Width = tbCtrlRightPane.Width
            dgDueToMe.Width = pnlDueToMeGrid.Width
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub pnlRightContainer_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlRightContainer.VisibleChanged

        If LoggedIn Then
            If Not bolpnlVisibilityChanged Then Exit Sub

            If pnlRightContainer.Visible Then
                cmbViewCalEntries.SelectedIndex = CInt(pnlRightContainer.Tag)
            Else
                pnlRightContainer.Tag = Nothing
                pnlRightContainer.Tag = CStr(cmbViewCalEntries.SelectedIndex)
            End If
            bolpnlVisibilityChanged = False
        End If

    End Sub
    Private Sub frmCalendarClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            If pCalendar.colIsDirty Then
                If (MsgBox("Do you want to Save Changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
                    pCalendar.Flush(CType(UIUtilsGen.ModuleID.[Global], Integer), MusterContainer.AppUser.UserKey, returnVal, AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If
            End If
            pCalendar.GetCalendarAll(IIf(cmbViewCalEntries.Text = String.Empty, AppUser.ID, cmbViewCalEntries.Text.ToString), False)

            chkToDoCompItems_CheckedChanged(sender, e)
            chkDueToMeCompItems_CheckedChanged(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub frmCalendarClosed(ByVal sender As Object, ByVal e As System.EventArgs)
    End Sub
    Private Sub calTechnicalMonth_DateSelected(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles calTechnicalMonth.DateSelected
        Dim tmpDate As Date

        Try

            tmpDate = e.Start
            tmpDate = New System.DateTime(tmpDate.Year, tmpDate.Month, 1, 0, 0, 0, 0)
            filterStartMonth = tmpDate.ToShortDateString.ToString()

            tmpDate = tmpDate.AddDays(tmpDate.DaysInMonth(tmpDate.Year, tmpDate.Month) - 1)
            filterEndMonth = tmpDate.ToShortDateString.ToString()
            tmpDate = e.Start
            tmpDate = tmpDate.AddDays(tmpDate.DayOfWeek.Sunday - tmpDate.DayOfWeek)
            filterStartWeek = tmpDate.ToShortDateString.ToString()
            tmpDate = tmpDate.AddDays(6)
            filterEndWeek = tmpDate.ToShortDateString.ToString()
            tmpDate = e.Start
            filterEndDay = tmpDate.ToShortDateString.ToString()
            setFilter()
            ''''''''''''''''''''''''''''''''''
            'userId or GroupID has to be passed as parameters at later time when integrated with login
            chkToDoCompItems_CheckedChanged(sender, e)
            chkDueToMeCompItems_CheckedChanged(sender, e)
            calendarRefreshTimer.Stop()
            calendarRefreshTimer.Start()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub chkToDoCompItems_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkToDoCompItems.CheckedChanged
        If chkToDoCompItems.Checked Then
            btnCalToDoModify.Enabled = False
            btnToDoMarkCompleted.Enabled = False
        Else
            btnCalToDoModify.Enabled = True
            btnToDoMarkCompleted.Enabled = True
        End If
        If bolCalLoad Then Exit Sub
        LoadToDoCalendar()
    End Sub

    Private Sub chkDueToMeCompItems_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDueToMeCompItems.CheckedChanged
        If chkDueToMeCompItems.Checked Then
            btnCalDueToMeModify.Enabled = False
            btnDueToMeMarkComp.Enabled = False
        Else
            btnCalDueToMeModify.Enabled = True
            btnDueToMeMarkComp.Enabled = True
        End If
        If bolCalLoad Then Exit Sub
        LoadDueToMeCalendar()
    End Sub
    Public Function RefreshCalendarInfo()
        pCalendar.GetCalendarAll(IIf(cmbViewCalEntries.Text = String.Empty, AppUser.ID, cmbViewCalEntries.Text.ToString), False)
        calendarRefreshTimer.Stop()
        calendarRefreshTimer.Start()
    End Function
    Public Function LoadDueToMeCalendar()

        Dim dtDueToMe As DataTable
        Try
            LockWindowUpdate(CLng(Me.Handle.ToInt64))
            dgDueToMe.DataSource = Nothing
            dtDueToMe = pCalendar.CalendarTableDueToMe()
            setFilter()
            If chkDueToMeCompItems.Checked Then
                dtDueToMe.DefaultView.RowFilter = completedTaskFilterString + " AND COMPLETED=TRUE"
                dtDueToMe.DefaultView.Sort = "DATE_DUE ASC"
                dgDueToMe.DataSource = dtDueToMe
                setDataGridWidth(dgDueToMe, 0)
            Else
                dtDueToMe.DefaultView.RowFilter = filterString + IIf(Not filterString Is Nothing, " AND COMPLETED=FALSE", "")
                dtDueToMe.DefaultView.Sort = "DATE_DUE ASC"
                dgDueToMe.DataSource = dtDueToMe
                setDataGridWidth(dgDueToMe, 0)
            End If
            dgDueToMe.DisplayLayout.Bands(0).Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            BoldCalendarDates(dtDueToMe)
            HideGridColumns(Me.dgDueToMe)
            ColorInOverDueEntries(dgDueToMe)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            LockWindowUpdate(0)
        End Try
    End Function

    Public Sub LoadToDoCalendar()

        Dim dtCalToDo As DataTable
        Dim dsCommon As New DataSet
        Try
            LockWindowUpdate(CLng(Me.Handle.ToInt64))
            dgCalToDo.DataSource = Nothing
            setFilter()
            dtCalToDo = pCalendar.CalendarTableToDo()
            dsCommon.Tables.Add(dtCalToDo)

            If chkToDoCompItems.Checked Then
                dtCalToDo.DefaultView.RowFilter = completedTaskFilterString + " AND COMPLETED=TRUE"
                dtCalToDo.DefaultView.Sort = "DATE_DUE ASC"
                Me.dgCalToDo.DataSource = dtCalToDo.DefaultView
                setDataGridWidth(dgCalToDo, 1)
            Else
                dtCalToDo.DefaultView.RowFilter = filterString + IIf(Not filterString Is Nothing, " AND COMPLETED=FALSE", "")
                dtCalToDo.DefaultView.Sort = "DATE_DUE ASC"
                Me.dgCalToDo.DataSource = dtCalToDo

                setDataGridWidth(dgCalToDo, 1)
            End If
            dgCalToDo.DisplayLayout.Bands(0).Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            BoldCalendarDates(dtCalToDo)
            HideGridColumns(Me.dgCalToDo)
            ColorInOverDueEntries(dgCalToDo)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            LockWindowUpdate(0)
        End Try
    End Sub

    Private Sub rdCalDay_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdCalDay.CheckedChanged
        Try
            If rdCalDay.Checked Then
                rOption = "D"
                setFilter()
                chkToDoCompItems_CheckedChanged(sender, e)
                chkDueToMeCompItems_CheckedChanged(sender, e)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub rdCalWeek_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdCalWeek.CheckedChanged
        Try
            If rdCalWeek.Checked Then
                rOption = "W"
                setFilter()
                chkToDoCompItems_CheckedChanged(sender, e)
                chkDueToMeCompItems_CheckedChanged(sender, e)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub rdCalMonth_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdCalMonth.CheckedChanged
        Try
            If rdCalMonth.Checked Then
                rOption = "M"
                setFilter()
                chkToDoCompItems_CheckedChanged(sender, e)
                chkDueToMeCompItems_CheckedChanged(sender, e)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub calTechnicalMonth_DateChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DateRangeEventArgs) Handles calTechnicalMonth.DateChanged
        If Not bolAllowDateChangedEvent Then Exit Sub
        If IsNothing(AppUser) Then
            Exit Sub
        ElseIf AppUser.ID = String.Empty Then
            Exit Sub
        End If
        Dim tmpDate As Date

        Try
            tmpDate = e.Start
            tmpDate = New System.DateTime(tmpDate.Year, tmpDate.Month, 1, 0, 0, 0, 0)
            filterStartMonth = tmpDate.ToShortDateString.ToString()
            tmpDate = tmpDate.AddDays(tmpDate.DaysInMonth(tmpDate.Year, tmpDate.Month) - 1)
            filterEndMonth = tmpDate.ToShortDateString.ToString()
            tmpDate = e.Start
            tmpDate = tmpDate.AddDays(tmpDate.DayOfWeek.Sunday - tmpDate.DayOfWeek)
            filterStartWeek = tmpDate.ToShortDateString.ToString()
            tmpDate = tmpDate.AddDays(6)
            filterEndWeek = tmpDate.ToShortDateString.ToString()
            tmpDate = e.Start
            filterEndDay = tmpDate.ToShortDateString.ToString()
            setFilter()
            ''''''''''''''''''''''''''''''''''
            'userId or GroupID has to be passed as parameters at later time when integrated with login
            chkToDoCompItems_CheckedChanged(sender, e)
            chkDueToMeCompItems_CheckedChanged(sender, e)
            calendarRefreshTimer.Stop()
            calendarRefreshTimer.Start()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbViewCalEntries_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbViewCalEntries.SelectedIndexChanged
        Try
            If bolCalLoad Then Exit Sub
            '''''''''''''''''''''''''''''''''' 
            'userId or GroupID has to be passed as parameters at later time when integrated with login
            If cmbViewCalEntries.SelectedIndex < 0 Then Exit Sub
            pCalendar.GetCalendarAll(cmbViewCalEntries.Text.ToString, False)
            chkToDoCompItems_CheckedChanged(sender, e)
            chkDueToMeCompItems_CheckedChanged(sender, e)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dgCalToDo_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles dgCalToDo.AfterSelectChange
        If ValidateUser(dgCalToDo) Then
            btnToDoDelete.Enabled = True
            btnCalToDoModify.Enabled = True
            btnToDoMarkCompleted.Enabled = True
        Else
            btnToDoDelete.Enabled = False
            btnCalToDoModify.Enabled = False
            btnToDoMarkCompleted.Enabled = False
            If AppUser.HEAD_ADMIN Then
                btnToDoMarkCompleted.Enabled = True
            End If
        End If

    End Sub
    Private Sub dgDueToMe_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles dgDueToMe.AfterSelectChange
        If ValidateUser(dgDueToMe) Then
            btnDueToMeDelete.Enabled = True
            btnCalDueToMeModify.Enabled = True
            btnDueToMeMarkComp.Enabled = True
        Else
            btnDueToMeDelete.Enabled = False
            btnCalDueToMeModify.Enabled = False
            btnDueToMeMarkComp.Enabled = False
            If AppUser.HEAD_ADMIN Then
                btnDueToMeMarkComp.Enabled = True
            End If
        End If
    End Sub
    Private Sub pnlRightContainer_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlRightContainer.Resize
        Dim nWidth As Integer

        If LoggedIn Then
            Try
                nWidth = calTechnicalMonth.Size.Width
                If nWidth > 192 Then
                    dgCalToDo.Size = New Size(nWidth, nGridHeight)
                    dgDueToMe.Size = New Size(nWidth, nGridHeight)
                    setDataGridWidth(dgCalToDo, 0)
                    setDataGridWidth(dgDueToMe, 1)
                    dgCalToDo.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
                    dgDueToMe.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
                Else
                    If dgCalToDo.DisplayLayout.Bands(0).Columns.Count > 0 And dgDueToMe.DisplayLayout.Bands(0).Columns.Count > 0 Then
                        dgCalToDo.DisplayLayout.AutoFitColumns = False
                        dgDueToMe.DisplayLayout.AutoFitColumns = False
                        dgCalToDo.Size = New Size(nGridWidth, nGridHeight)
                        dgDueToMe.Size = New Size(nGridWidth, nGridHeight)
                        setDataGridWidth(dgCalToDo, 0)
                        setDataGridWidth(dgDueToMe, 1)
                        dgCalToDo.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Automatic
                        dgDueToMe.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Automatic
                    End If
                End If
            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            End Try
        End If

    End Sub

    Private Sub dgCalToDo_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgCalToDo.DoubleClick
        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        CalendarDoubleClick(dgCalToDo)
    End Sub
    Private Sub dgDueToMe_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgDueToMe.DoubleClick
        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        CalendarDoubleClick(dgDueToMe)
    End Sub
    Private Sub CalendarDoubleClick(ByVal ugCal As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim bolSuccess As Boolean = False
        Try
            If Not ugCal.ActiveRow Is Nothing Then
                Dim fromEntityType As Int64 = ugCal.ActiveRow.Cells("Owning_Entity_Type").Value
                Dim toEntityID As Int64 = 0
                Select Case fromEntityType
                    Case UIUtilsGen.EntityTypes.LUST_Event, _
                            UIUtilsGen.EntityTypes.LustActivity, _
                            UIUtilsGen.EntityTypes.LustDocument
                        ' only if user has read/write access to the module
                        If cmbSearchModule.FindString(UIUtilsGen.ModuleID.Technical.ToString) > 0 Then
                            toEntityID = pCalendar.GetParentEntityID(ugCal.ActiveRow.Cells("Owning_Entity_ID").Value, fromEntityType, UIUtilsGen.EntityTypes.Facility)
                            If toEntityID <> 0 Then
                                txtOwnerQSKeyword.Text = toEntityID
                                cmbSearchModule.SelectedIndex = cmbSearchModule.FindString(UIUtilsGen.ModuleID.Technical.ToString)
                                cmbQuickSearchFilter.SelectedIndex = cmbQuickSearchFilter.FindString("Facility ID")
                                bolSuccess = True
                            End If
                        Else
                            MsgBox("Entity Type: " + fromEntityType.ToString + vbCrLf + _
                                    "Task Desc: " + ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text + vbCrLf + _
                                    "Group ID: " + ugCal.ActiveRow.Cells("GROUP_ID").Text + vbCrLf + _
                                    "not handled. Please notify Administrator")
                        End If
                    Case UIUtilsGen.EntityTypes.Tank
                        ' only if user has read/write access to the module
                        If cmbSearchModule.FindString(UIUtilsGen.ModuleID.Registration.ToString) > 0 Then
                            toEntityID = pCalendar.GetParentEntityID(ugCal.ActiveRow.Cells("Owning_Entity_ID").Value, fromEntityType, UIUtilsGen.EntityTypes.Facility)
                            If toEntityID <> 0 Then
                                txtOwnerQSKeyword.Text = toEntityID
                                cmbSearchModule.SelectedIndex = cmbSearchModule.FindString(UIUtilsGen.ModuleID.Registration.ToString)
                                cmbQuickSearchFilter.SelectedIndex = cmbQuickSearchFilter.FindString("Facility ID")
                                bolSuccess = True
                            End If
                        Else
                            MsgBox("Entity Type: " + fromEntityType.ToString + vbCrLf + _
                                    "Task Desc: " + ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text + vbCrLf + _
                                    "Group ID: " + ugCal.ActiveRow.Cells("GROUP_ID").Text + vbCrLf + _
                                    "not handled. Please notify Administrator")
                        End If
                    Case UIUtilsGen.EntityTypes.ClosureEvent
                        ' only if user has read/write access to the module
                        If cmbSearchModule.FindString(UIUtilsGen.ModuleID.Closure.ToString) > 0 Then
                            toEntityID = pCalendar.GetParentEntityID(ugCal.ActiveRow.Cells("Owning_Entity_ID").Value, fromEntityType, UIUtilsGen.EntityTypes.Facility)
                            If toEntityID <> 0 Then
                                txtOwnerQSKeyword.Text = toEntityID
                                cmbSearchModule.SelectedIndex = cmbSearchModule.FindString(UIUtilsGen.ModuleID.Closure.ToString)
                                cmbQuickSearchFilter.SelectedIndex = cmbQuickSearchFilter.FindString("Facility ID")
                                bolSuccess = True
                            End If
                        Else
                            MsgBox("Entity Type: " + fromEntityType.ToString + vbCrLf + _
                                    "Task Desc: " + ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text + vbCrLf + _
                                    "Group ID: " + ugCal.ActiveRow.Cells("GROUP_ID").Text + vbCrLf + _
                                    "not handled. Please notify Administrator")
                        End If
                    Case UIUtilsGen.EntityTypes.Fees
                        ' only if user has read/write access to the module
                        If cmbSearchModule.FindString(UIUtilsGen.ModuleID.Fees.ToString) > 0 Then
                            toEntityID = pCalendar.GetParentEntityID(ugCal.ActiveRow.Cells("Owning_Entity_ID").Value, fromEntityType, UIUtilsGen.EntityTypes.Owner)
                            If toEntityID <> 0 Then
                                txtOwnerQSKeyword.Text = toEntityID
                                cmbSearchModule.SelectedIndex = cmbSearchModule.FindString(UIUtilsGen.ModuleID.Fees.ToString)
                                cmbQuickSearchFilter.SelectedIndex = cmbQuickSearchFilter.FindString("Owner ID")
                                bolSuccess = True
                            End If
                        Else
                            MsgBox("Entity Type: " + fromEntityType.ToString + vbCrLf + _
                                    "Task Desc: " + ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text + vbCrLf + _
                                    "Group ID: " + ugCal.ActiveRow.Cells("GROUP_ID").Text + vbCrLf + _
                                    "not handled. Please notify Administrator")
                        End If
                    Case UIUtilsGen.EntityTypes.Facility
                        If ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text.IndexOf("incomplete registration") > -1 Or _
                            ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text.IndexOf("Signature Required Letter") > -1 Then
                            ' only if user has read/write access to the module
                            If cmbSearchModule.FindString(UIUtilsGen.ModuleID.Registration.ToString) > 0 Then
                                toEntityID = ugCal.ActiveRow.Cells("Owning_Entity_ID").Value
                                If toEntityID <> 0 Then
                                    txtOwnerQSKeyword.Text = toEntityID
                                    cmbSearchModule.SelectedIndex = cmbSearchModule.FindString(UIUtilsGen.ModuleID.Registration.ToString)
                                    cmbQuickSearchFilter.SelectedIndex = cmbQuickSearchFilter.FindString("Facility ID")
                                    bolSuccess = True
                                End If
                            End If
                        ElseIf Not ugCal.ActiveRow.Cells("GROUP_ID").Value Is DBNull.Value Then
                            If cmbSearchModule.FindString(ugCal.ActiveRow.Cells("GROUP_ID").Text) > 0 Then
                                toEntityID = ugCal.ActiveRow.Cells("Owning_Entity_ID").Value
                                If toEntityID <> 0 Then
                                    txtOwnerQSKeyword.Text = toEntityID
                                    cmbSearchModule.SelectedIndex = cmbSearchModule.FindString(ugCal.ActiveRow.Cells("GROUP_ID").Text)
                                    cmbQuickSearchFilter.SelectedIndex = cmbQuickSearchFilter.FindString("Facility ID")
                                    bolSuccess = True
                                End If
                            End If
                        Else
                            MsgBox("Entity Type: " + fromEntityType.ToString + vbCrLf + _
                                    "Task Desc: " + ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text + vbCrLf + _
                                    "Group ID: " + ugCal.ActiveRow.Cells("GROUP_ID").Text + vbCrLf + _
                                    "not handled. Please notify Administrator")
                        End If
                    Case UIUtilsGen.EntityTypes.Owner
                        If ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text.IndexOf("incomplete registration") > -1 Then
                            ' only if user has read/write access to the module
                            If cmbSearchModule.FindString(UIUtilsGen.ModuleID.Registration.ToString) > 0 Then
                                toEntityID = ugCal.ActiveRow.Cells("Owning_Entity_ID").Value
                                If toEntityID <> 0 Then
                                    txtOwnerQSKeyword.Text = toEntityID
                                    cmbSearchModule.SelectedIndex = cmbSearchModule.FindString(UIUtilsGen.ModuleID.Registration.ToString)
                                    cmbQuickSearchFilter.SelectedIndex = cmbQuickSearchFilter.FindString("Owner ID")
                                    bolSuccess = True
                                End If
                            End If
                        ElseIf Not ugCal.ActiveRow.Cells("GROUP_ID").Value Is DBNull.Value Then
                            If cmbSearchModule.FindString(ugCal.ActiveRow.Cells("GROUP_ID").Text) > 0 Then
                                toEntityID = ugCal.ActiveRow.Cells("Owning_Entity_ID").Value
                                If toEntityID <> 0 Then
                                    txtOwnerQSKeyword.Text = toEntityID
                                    cmbSearchModule.SelectedIndex = cmbSearchModule.FindString(ugCal.ActiveRow.Cells("GROUP_ID").Text)
                                    cmbQuickSearchFilter.SelectedIndex = cmbQuickSearchFilter.FindString("Owner ID")
                                    bolSuccess = True
                                End If
                            End If
                        Else
                            MsgBox("Entity Type: " + fromEntityType.ToString + vbCrLf + _
                                    "Task Desc: " + ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text + vbCrLf + _
                                    "Group ID: " + ugCal.ActiveRow.Cells("GROUP_ID").Text + vbCrLf + _
                                    "not handled. Please notify Administrator")
                        End If
                    Case Else
                        MsgBox("Entity Type: " + fromEntityType.ToString + vbCrLf + _
                                "Task Desc: " + ugCal.ActiveRow.Cells("TASK_DESCRIPTION").Text + vbCrLf + _
                                "Group ID: " + ugCal.ActiveRow.Cells("GROUP_ID").Text + vbCrLf + _
                                "not handled. Please notify Administrator")
                End Select
            End If
            If bolSuccess Then
                btnQuickOwnerSearch.PerformClick()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub calendarRefreshTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles calendarRefreshTimer.Tick
        If bolAllowDateChangedEvent Then Exit Sub
        Try
            bolAllowDateChangedEvent = True
            Dim ea As New System.windows.Forms.DateRangeEventArgs(filterStartWeek, filterEndWeek)
            calTechnicalMonth_DateChanged(calTechnicalMonth, ea)
        Finally
            bolAllowDateChangedEvent = False
        End Try
    End Sub
#End Region

#Region "Other Functions"
    Private Sub ColumnsToHide(ByVal nIndex As Integer)
        Dim i As Integer
        Try
            For i = 3 To 9
                dsCommon.Tables(nIndex).Columns(i).ColumnMapping = MappingType.Hidden
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    'Public Sub SetEnableValues(ByVal sender As Object, ByVal e As DataGridEnableEventArgs)

    'End Sub
    'Private Function loadTasks(ByVal calTask As InfoRepository.CalendarTask, Optional ByVal wrkFlow As String = Nothing) As DataSet
    '    Dim dsCal As DataSet
    '    Dim CalConsumer As New CalendarConsumer
    '    Try
    '        dsCal = CalConsumer.getTasks(calTask, wrkFlow)
    '        Return dsCal
    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Function

#End Region

#End Region

#Region "Barometer"
    Private Sub LoadBarometer(Optional ByVal entityID As Integer = -1, Optional ByVal entityType As Integer = -1, Optional ByVal [Module] As String = "NONE", Optional ByVal ParentFormText As String = "NONE", Optional ByVal eventID As Integer = -1, Optional ByVal eventType As Integer = -1)
        Try
            If entityID = -1 Then entityID = flagBarometerEntityID
            If entityType = -1 Then entityType = flagBarometerEntityType
            If [Module] = "NONE" Then [Module] = flagBarometerModule
            If ParentFormText = "NONE" Then ParentFormText = flagBarometerParentFormText
            If eventID = -1 Then eventID = 0
            If eventType = -1 Then eventType = 0

            Select Case [Module]
                Case "Registration", "Closure", "Technical", "Financial", "C & E", "Company", "Licensee", "Fees"
                    Dim ds As DataSet
                    ds = pFlag.GetBarometerColors(entityID, entityType, eventID, eventType)

                    ClearBarometer()

                    flagBarometerEntityID = entityID
                    flagBarometerEntityType = entityType
                    flagBarometerModule = [Module]
                    flagBarometerParentFormText = ParentFormText
                    flagBarometerEventID = eventID
                    flagBarometerEventType = eventType

                    If ds.Tables.Count > 0 Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            For Each col As DataColumn In ds.Tables(0).Columns
                                Select Case col.ColumnName
                                    Case "OWNER"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnOwnerFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnOwnerFlag.BackColor = Color.Yellow
                                        End If
                                    Case "FACILITY"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnFacFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnFacFlag.BackColor = Color.Yellow
                                        End If
                                    Case "FEES"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnFeeFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnFeeFlag.BackColor = Color.Yellow
                                        End If
                                    Case "CLOSURE"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnCloFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnCloFlag.BackColor = Color.Yellow
                                        End If
                                    Case "LUST"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnLUSFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnLUSFlag.BackColor = Color.Yellow
                                        End If
                                    Case "FIN"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnFinFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnFinFlag.BackColor = Color.Yellow
                                        End If
                                    Case "INSP"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnInsFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnInsFlag.BackColor = Color.Yellow
                                        End If
                                    Case "CAE"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnCandEFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnCandEFlag.BackColor = Color.Yellow
                                        End If
                                    Case "FIRM"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnFirmLicFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnFirmLicFlag.BackColor = Color.Yellow
                                        End If
                                    Case "IND"
                                        If ds.Tables(0).Rows(0)(col.ColumnName) = "RED" Then
                                            btnIndvLicFlag.BackColor = Color.Red
                                        ElseIf ds.Tables(0).Rows(0)(col.ColumnName) = "YELLOW" Then
                                            btnIndvLicFlag.BackColor = Color.Yellow
                                        End If
                                End Select
                            Next
                        End If
                    End If

                    Select Case flagBarometerModule
                        Case "Registration"
                            btnOwnerFlag.Visible = True
                            btnFacFlag.Visible = True
                            btnFeeFlag.Visible = True
                            btnCloFlag.Visible = True
                            btnLUSFlag.Visible = True
                            btnFinFlag.Visible = True
                            btnCandEFlag.Visible = True
                            btnInsFlag.Visible = True
                            btnIndvLicFlag.Visible = True
                        Case "Closure"
                            btnOwnerFlag.Visible = True
                            btnFacFlag.Visible = True
                            btnFeeFlag.Visible = True
                            btnCloFlag.Visible = True
                            btnLUSFlag.Visible = True
                            btnFinFlag.Visible = True
                            btnCandEFlag.Visible = True
                            btnInsFlag.Visible = True
                            btnFirmLicFlag.Visible = True
                            btnIndvLicFlag.Visible = True
                        Case "Technical"
                            btnOwnerFlag.Visible = True
                            btnFacFlag.Visible = True
                            btnFeeFlag.Visible = True
                            btnCloFlag.Visible = True
                            btnLUSFlag.Visible = True
                            btnFinFlag.Visible = True
                            btnCandEFlag.Visible = True
                            btnInsFlag.Visible = True
                            btnFirmLicFlag.Visible = True
                            btnIndvLicFlag.Visible = True
                        Case "Financial"
                            btnOwnerFlag.Visible = True
                            btnFacFlag.Visible = True
                            btnFeeFlag.Visible = True
                            btnCloFlag.Visible = True
                            btnLUSFlag.Visible = True
                            btnFinFlag.Visible = True
                            btnCandEFlag.Visible = True
                            btnInsFlag.Visible = True
                            btnFirmLicFlag.Visible = True
                            btnIndvLicFlag.Visible = True
                            'Case "Inspection"
                            '    btnOwnerFlag.Visible = True
                            '    btnFacFlag.Visible = True
                            '    btnFeeFlag.Visible = True
                            '    btnCloFlag.Visible = True
                            '    btnLUSFlag.Visible = True
                            '    btnFinFlag.Visible = True
                            '    btnCandEFlag.Visible = True
                            '    btnInsFlag.Visible = True
                            '    btnFirmLicFlag.Visible = True
                            '    btnIndvLicFlag.Visible = True
                        Case "C & E"
                            btnOwnerFlag.Visible = True
                            btnFacFlag.Visible = True
                            btnFeeFlag.Visible = True
                            btnCloFlag.Visible = True
                            btnLUSFlag.Visible = True
                            btnFinFlag.Visible = True
                            btnCandEFlag.Visible = True
                            btnInsFlag.Visible = True
                            btnFirmLicFlag.Visible = True
                            btnIndvLicFlag.Visible = True
                        Case "Company"
                            btnFirmLicFlag.Visible = True
                            btnIndvLicFlag.Visible = True
                        Case "Licensee"
                            btnCandEFlag.Visible = True
                        Case "Fees"
                            btnOwnerFlag.Visible = True
                            btnFacFlag.Visible = True
                            btnFeeFlag.Visible = True
                            btnCloFlag.Visible = True
                            btnLUSFlag.Visible = True
                            btnFinFlag.Visible = True
                            btnCandEFlag.Visible = True
                            btnInsFlag.Visible = True
                            btnFirmLicFlag.Visible = True
                            btnIndvLicFlag.Visible = True
                            'Case Else
                            '    btnOwnerFlag.Visible = True
                            '    btnFacFlag.Visible = True
                            '    btnFeeFlag.Visible = True
                            '    btnCloFlag.Visible = True
                            '    btnLUSFlag.Visible = True
                            '    btnFinFlag.Visible = True
                            '    btnCandEFlag.Visible = True
                            '    btnInsFlag.Visible = True
                            '    btnFirmLicFlag.Visible = True
                            '    btnIndvLicFlag.Visible = True
                    End Select
                Case Else
                    ClearBarometer()
            End Select

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ClearBarometer()
        Try
            btnOwnerFlag.BackColor = System.Drawing.SystemColors.Control
            btnFacFlag.BackColor = System.Drawing.SystemColors.Control
            btnFeeFlag.BackColor = System.Drawing.SystemColors.Control
            btnCloFlag.BackColor = System.Drawing.SystemColors.Control
            btnLUSFlag.BackColor = System.Drawing.SystemColors.Control
            btnFinFlag.BackColor = System.Drawing.SystemColors.Control
            btnCandEFlag.BackColor = System.Drawing.SystemColors.Control
            btnInsFlag.BackColor = System.Drawing.SystemColors.Control
            btnFirmLicFlag.BackColor = System.Drawing.SystemColors.Control
            btnIndvLicFlag.BackColor = System.Drawing.SystemColors.Control

            btnOwnerFlag.Visible = False
            btnFacFlag.Visible = False
            btnFeeFlag.Visible = False
            btnCloFlag.Visible = False
            btnLUSFlag.Visible = False
            btnFinFlag.Visible = False
            btnCandEFlag.Visible = False
            btnInsFlag.Visible = False
            btnFirmLicFlag.Visible = False
            btnIndvLicFlag.Visible = False

            flagBarometerEntityID = 0
            flagBarometerEntityType = 0
            flagBarometerModule = String.Empty
            flagBarometerParentFormText = String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ShowFlags(ByVal btn As Button)
        Try
            If flagBarometerEntityID = 0 And flagBarometerEntityType = 0 And _
                flagBarometerModule = String.Empty And _
                flagBarometerParentFormText = String.Empty Then
                Exit Sub
            End If

            If flagBarometerEntityType = 9 AndAlso flagBarometerModule = "Registration" Then
                flagBarometerModule = "All"
            End If

            Select Case btn.Name
                Case btnOwnerFlag.Name
                    SF = New ShowFlags(flagBarometerOwnerID, UIUtilsGen.EntityTypes.Owner, flagBarometerModule, , , , , , Me)
                Case btnFacFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "Registration", True, , , , , Me)
                Case btnFeeFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "Fees", True, , , , , Me)
                Case btnCloFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "Closure", True, , , , , Me)
                Case btnLUSFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "Technical", True, , , , , Me)
                Case btnFinFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "Financial", True, , , , , Me)
                Case btnCandEFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "C & E", True, , , , , Me)
                Case btnInsFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "Inspection", True, , , , , Me)
                Case btnFirmLicFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "Company", True, , , , , Me)
                Case btnIndvLicFlag.Name
                    SF = New ShowFlags(flagBarometerEntityID, flagBarometerEntityType, "Licensee", True, , , , , Me)
            End Select
            If Not (SF Is Nothing) Then
                SF.ShowDialog()
                SF = Nothing
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerFlag.Click
        Try
            ShowFlags(btnOwnerFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFacFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacFlag.Click
        Try
            ShowFlags(btnFacFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFeeFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFeeFlag.Click
        Try
            ShowFlags(btnFeeFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCloFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloFlag.Click
        Try
            ShowFlags(btnCloFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLUSFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLUSFlag.Click
        Try
            ShowFlags(btnLUSFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFinFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinFlag.Click
        Try
            ShowFlags(btnFinFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCandEFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCandEFlag.Click
        Try
            ShowFlags(btnCandEFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnInsFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsFlag.Click
        Try
            ShowFlags(btnInsFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFirmLicFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirmLicFlag.Click
        Try
            ShowFlags(btnFirmLicFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnIndvLicFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIndvLicFlag.Click
        Try
            ShowFlags(btnIndvLicFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

    Private Sub CheckMouseToVisibleCalendar(Optional ByVal force As Boolean = False)


        If Me.pnlRightContainer.Visible Then

            Try


                If Cursor.Position.X < pnlRightContainer.PointToScreen(New Point(0, 0)).X OrElse _
                   Cursor.Position.Y < pnlRightContainer.PointToScreen(New Point(0, 0)).Y OrElse _
                   Cursor.Position.Y > pnlRightContainer.PointToScreen(New Point(0, pnlRightContainer.Height)).Y OrElse _
                   force Then

                    Me.pnlRightMost.Size = New System.Drawing.Size(16, 701)
                    lockChildren()
                    bolpnlVisibilityChanged = True
                    btnDock.Text = "<"
                    pnlRightContainer.Visible = False
                    RightPanelDockFlag = False
                End If

            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            Finally
                releaseChildren()
            End Try
        End If

    End Sub



    Private Sub CheckMenuItemRights(ByRef dtTable As DataTable)
        Try

            mnuItemRegistration.Enabled = False
            mnItemCAE.Enabled = False
            mnItemTechnical.Enabled = False
            mnuItemInspector.Enabled = False
            mnuItemFinancial.Enabled = False
            mnuItemClosure.Enabled = False
            mnuItemFees.Enabled = False
            mnuItemCompanyModule.Enabled = False
            mnuItmContact.Enabled = False
            mnuSubItemAdminServices.Enabled = False
            mnItemCAP.Enabled = False

            'If AppUser.HEAD_CANDE Then
            '    mnItemCAE.Enabled = True
            'End If
            If AppUser.HEAD_FEES Then
                mnuSubItemAdminServices.Enabled = True
            End If

            Dim drow As DataRow
            For Each drow In dtTable.Rows
                If drow(0) = UIUtilsGen.ModuleID.Registration Then
                    mnuItemRegistration.Enabled = True
                    'ElseIf drow(0) = 613 Then
                    'mnItemCAE.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.CAPProcess Then
                    mnItemCAP.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.Technical Then
                    mnItemTechnical.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.Inspection Then
                    mnuItemInspector.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.Financial Then
                    mnuItemFinancial.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.Closure Then
                    mnuItemClosure.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.Fees Then
                    mnuItemFees.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.FeeAdmin Then
                    mnuSubItemAdminServices.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.Company Then
                    mnuItemCompanyModule.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.ContactManagement Then
                    mnuItmContact.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.CAE Then
                    mnItemCAE.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.Admin Then
                    AddAdmin()
                    mnuSubItemAdminServices.Enabled = True
                    mnItemCAE.Enabled = True
                    mnItemCAP.Enabled = True
                ElseIf drow(0) = UIUtilsGen.ModuleID.TechAdmin Then
                    AddTechAdmin()
                ElseIf drow(0) = UIUtilsGen.ModuleID.FinAdmin Then
                    AddFinAdmin()
                ElseIf drow(0) = UIUtilsGen.ModuleID.CompanyAdmin Then
                    AddCompanyAdmin()
                End If
            Next

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Friend Shared Function GetWordApp() As Word.Application
        Dim boltest As Boolean
        Try


            For Each p As Process In Process.GetProcessesByName("winword")

                If p.MainWindowHandle.Equals(System.IntPtr.Zero) Then
                    p.Kill()
                    WordApp = Nothing
                End If

            Next
        Catch
        End Try


        Try

            If IsNothing(WordApp) Then
                WordApp = GetObject(, "Word.Application")
            End If


            boltest = WordApp.Visible

        Catch ex As Exception
            Try
                If ex.Message.ToUpper = "Cannot Create ActiveX Component.".ToUpper Then
                    WordApp = New Word.Application
                ElseIf ex.Message.ToUpper = "The RPC server is unavailable.".ToUpper Then
                    WordApp = New Word.Application
                Else
                    Try
                        For Each p As Process In Process.GetProcessesByName("winword")

                            If p.MainWindowHandle.Equals(System.IntPtr.Zero) Then
                                p.Kill()
                                WordApp = Nothing
                            End If

                            WordApp = GetObject(, "Word.Application")

                        Next
                    Catch ex3 As Exception
                        Throw ex
                    End Try
                End If
            Catch ex2 As Exception
                MsgBox("Document cannot be created. Please make sure Microsoft word is installed and any unused Word applications are opened", MsgBoxStyle.OKOnly, "Word Document Interface")

                Return Nothing
            End Try

        End Try
        Return WordApp
    End Function

    Private Sub HideShowOwnerInfo(ByVal visibility As Boolean)
        Try
            lblFacility.Visible = visibility
            lblFacilityAddress.Visible = visibility
            lblOwner.Visible = visibility
            lblOwnerName.Visible = visibility
            lblOwnerAddress.Visible = True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub OnCalModeClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    MsgBox("Cal Mode")
    'End Sub
    'Private Sub OnSearchModeClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    MsgBox("Search Mode")
    'End Sub
    Private Sub FormToDisplay(ByVal strCurrentForm As String)
        Dim frm As Form
        Dim frmTag As String
        For Each frm In Me.MdiChildren
            frmTag = CStr(frm.Tag)
            If frmTag = strCurrentForm Then
                objRegister = frm
                objRegister.Show()
                objRegister.BringToFront()
                If frmTag = "Owner-Add" Then
                    objRegister.tbCntrlRegistration.SelectedTab = objRegister.tbPageOwnerDetail
                End If
                Exit Sub
            End If
        Next
    End Sub
    Private Sub SwitchToWindow(ByVal WindowTitle As String)
        Dim frm As Form
        For Each frm In Me.MdiChildren
            If frm.Text = WindowTitle Then
                frm.Activate()
                frm.Size = Me.ClientSize
                frm.BringToFront()
                Exit For
            End If
        Next
    End Sub

    Private Sub NavigateChildren(ByVal Direction As String)

        Dim ChildForm As Form
        Dim ChildArr As New ArrayList
        Dim CurrIndex As Int16
        Dim frm As Windows.Forms.Form
        Dim strTitle As String
        Dim ThisGUID As System.Guid
        '
        '  Get the title of the active window in the MDI container
        Try
            ThisGUID = AppSemaphores.GetValuePair("0", "ActiveForm")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                Dim MyErr As ErrorReport
                MyErr = New ErrorReport(New Exception("No windows open. You must first open a window.", ex))
                MyErr.ShowDialog()
                Exit Sub
            End If
        End Try

        Try
            strTitle = AppSemaphores.GetValuePair(ThisGUID.ToString, "WindowName")
        Catch ex As Exception
            If ex.Message.StartsWith("Argument 'Index'") Then
                Dim MyErr As ErrorReport
                MyErr = New ErrorReport(New Exception("No windows open.  You must first open a window.", ex))
                MyErr.ShowDialog()
                Exit Sub
            End If
        End Try

        '  Now build the list of active forms (excluding owner search)
        '
        For Each ChildForm In Me.MdiChildren
            ChildArr.Add(ChildForm)
            If ChildForm.Text = strTitle Then
                CurrIndex = ChildArr.Count - 1
            End If
        Next
        '
        ' Determine the index (in the array list) for the form in the indicated direction
        '
        Select Case Direction
            Case "Next"
                If CurrIndex + 1 > ChildArr.Count - 1 Then
                    CurrIndex = 0
                Else
                    CurrIndex += 1
                End If
            Case "Previous"
                If CurrIndex = 0 Then
                    CurrIndex = ChildArr.Count - 1
                Else
                    CurrIndex -= 1
                End If
        End Select

        '
        '  Activate the form
        '

        frm = CType(ChildArr(CurrIndex), Form)
        frm.Activate()
        frm.BringToFront()

    End Sub
    Private Sub EnableDisableControls(ByVal WindowTitle As String)
        Dim frm As Form
        Dim winTitle As String

        Try
            If WindowTitle.StartsWith("Registration - Owner Detail ()") Then 'And objRegister.lblOwnerIDValue.Text <> String.Empty Then

            ElseIf WindowTitle.StartsWith("Registration - Facility Detail ()") Then 'And objRegister.lblFacilityIDValue.Text <> String.Empty Then

            ElseIf WindowTitle.StartsWith("Registration - Manage Tank ()") Then 'And objRegister.lblTankIDValue.Text <> String.Empty Then

            ElseIf WindowTitle.StartsWith("Registration - Manage Pipe ()") Then 'And objRegister.lblPipeIDValue.Text <> String.Empty Then

            Else

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub RefreshHeaderInfo()
        Dim MyValue As String
        Dim x As New DateTime
        Dim entityID, entityType, eventID, eventType As Integer
        eventID = -1
        eventType = -1
        Try
            If Me.ActiveMdiChild Is Nothing Then Exit Sub

            If Me.ActiveMdiChild.Name.ToUpper.IndexOf("REGISTRATION") > -1 Then
                strModuleTitle = "Registration"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("CANDEMANAGEMENT") > -1 Then
                strModuleTitle = "C&&E Admin"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("INSPECTORS") > -1 Then
                strModuleTitle = "C&&E Admin"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("ADMIN") > -1 Or Me.ActiveMdiChild.Name.ToUpper.IndexOf("MANAGE") > -1 Then
                strModuleTitle = "Admin"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("TECHNICAL") > -1 Then
                strModuleTitle = "LUST"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("CLOSURE") > -1 Then
                strModuleTitle = "Closure"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("FINANCIAL") > -1 Then
                strModuleTitle = "Financial"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("INSPECTION") > -1 Then
                strModuleTitle = "Inspection"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("FEES") > -1 Then
                strModuleTitle = "Fees"
            ElseIf Me.ActiveMdiChild.Name.ToUpper.IndexOf("C " + "&" + " E") > -1 Or Me.ActiveMdiChild.Name.ToUpper.IndexOf("CANDE") > -1 Then
                strModuleTitle = "C " + "&" + " E"
            ElseIf Me.ActiveMdiChild.Name.ToUpper = "COMPANY" Then
                strModuleTitle = "Company"
            ElseIf Me.ActiveMdiChild.Name.ToUpper = "CAPPREMONTHLY" Then
                strModuleTitle = "CAP"
            Else
                strModuleTitle = String.Empty
            End If

            Me.pnlCommonReferenceArea.Visible = False
            Me.lblModuleName.Visible = False
            Me.lblOwner.Visible = False
            Me.lnklblPrevForm.Visible = False
            Me.lblOwnerInfo.Visible = False
            Me.lblFacility.Visible = False
            Me.lnkLblNextForm.Visible = False
            Me.lblFacilityID.Visible = False
            Me.lblFacilityInfo.Visible = False
            Me.lblOwner.Text = "Owner"
            If strModuleTitle <> String.Empty Then
                Dim bolGetSystemColor As Boolean = False
                If AppProfileInfo.Retrieve(AppUser.ID & "|MODULE_ID|" & strModuleTitle.ToUpper & "|HEADERCOLOR") Is Nothing Then
                    bolGetSystemColor = True
                ElseIf AppProfileInfo.ProfileValue Is Nothing Then
                    bolGetSystemColor = True
                End If
                If bolGetSystemColor Then
                    If AppProfileInfo.Retrieve("SYSTEM|MODULE_ID|" & strModuleTitle.ToUpper & "|HEADERCOLOR") Is Nothing Then
                        MyValue = System.Drawing.Color.FromName("LimeGreen").ToArgb.ToString
                    ElseIf AppProfileInfo.ProfileValue Is Nothing Then
                        MyValue = System.Drawing.Color.FromName("LimeGreen").ToArgb.ToString
                    Else
                        MyValue = AppProfileInfo.ProfileValue
                    End If
                    AppProfileInfo.Add(New MUSTER.Info.ProfileInfo(AppUser.ID, "MODULE_ID", _
                                            strModuleTitle.ToUpper, "HEADERCOLOR", MyValue, _
                                            False, _
                                            AppUser.ID, _
                                            Now, _
                                            String.Empty, _
                                            CDate("01/01/0001")))

                    AppProfileInfo.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), AppUser.UserKey, returnVal, True)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If

                If IsNumeric(AppProfileInfo.ProfileValue) Then
                    Me.pnlCommonReferenceArea.BackColor = System.Drawing.Color.FromArgb(AppProfileInfo.ProfileValue)
                Else
                    Me.pnlCommonReferenceArea.BackColor = System.Drawing.Color.FromArgb(CInt("&H" & AppProfileInfo.ProfileValue.Substring(0, 2)), CInt("&H" & AppProfileInfo.ProfileValue.Substring(2, 2)), CInt("&H" & AppProfileInfo.ProfileValue.Substring(4, 2)), CInt("&H" & AppProfileInfo.ProfileValue.Substring(6, 2)))
                End If

                If strModuleTitle = "C " + "&" + " E" Then
                    lblModuleName.Text = "C && E"
                Else
                    lblModuleName.Text = strModuleTitle
                End If
                Me.lblModuleName.Visible = True
                Me.pnlCommonReferenceArea.Visible = True
                If strModuleTitle.ToUpper = "REGISTRATION" Then
                    Me.lblOwner.Visible = True
                    Me.lnklblPrevForm.Visible = True
                    Me.lblOwnerInfo.Visible = True
                    Me.lblFacility.Visible = True
                    Me.lnkLblNextForm.Visible = True
                    Me.lblFacilityID.Visible = True
                    Me.lblFacilityInfo.Visible = True
                    Dim regLocal As Registration = CType(ActiveMdiChild, Registration)
                    entityID = IIf(regLocal.lblOwnerIDValue.Text.Trim = String.Empty, 0, regLocal.lblOwnerIDValue.Text.Trim)
                    entityType = UIUtilsGen.EntityTypes.Owner
                    flagBarometerOwnerID = entityID
                    If regLocal.lblFacilityIDValue.Text.Trim <> String.Empty Then
                        entityID = CType(regLocal.lblFacilityIDValue.Text.Trim, Integer)
                        entityType = UIUtilsGen.EntityTypes.Facility
                        If regLocal.nPipeID > 0 Then
                            eventID = regLocal.nPipeID
                            eventType = UIUtilsGen.EntityTypes.Pipe
                        ElseIf regLocal.nTankID > 0 Then
                            eventID = regLocal.nTankID
                            eventType = UIUtilsGen.EntityTypes.Tank
                        End If
                        flagBarometerOwnerID = CType(regLocal.lblOwnerIDValue.Text, Integer)
                    End If
                ElseIf strModuleTitle.ToUpper = "LUST" Then
                    Me.lblOwner.Visible = True
                    Me.lblOwnerInfo.Visible = True
                    Me.lblFacility.Visible = True
                    Me.lblFacilityID.Visible = True
                    Me.lblFacilityInfo.Visible = True
                    Dim techLocal As Technical = CType(ActiveMdiChild, Technical)
                    entityID = IIf(techLocal.lblOwnerIDValue.Text.Trim = String.Empty, 0, CType(techLocal.lblOwnerIDValue.Text.Trim, Integer))
                    entityType = UIUtilsGen.EntityTypes.Owner
                    flagBarometerOwnerID = entityID
                    If techLocal.lblFacilityIDValue.Text.Trim <> String.Empty Then
                        entityID = CType(techLocal.lblFacilityIDValue.Text.Trim, Integer)
                        entityType = UIUtilsGen.EntityTypes.Facility
                        If techLocal.nCurrentEventID > 0 Or techLocal.nCurrentEventID < -100 Then
                            eventID = techLocal.nCurrentEventID
                            eventType = UIUtilsGen.EntityTypes.LUST_Event
                        End If
                    End If
                ElseIf strModuleTitle.ToUpper = "CLOSURE" Then
                    Me.lblOwner.Visible = True
                    Me.lblOwnerInfo.Visible = True
                    Me.lblFacility.Visible = True
                    Me.lblFacilityID.Visible = True
                    Me.lblFacilityInfo.Visible = True
                    Dim cloLocal As Closure = CType(ActiveMdiChild, Closure)
                    entityID = IIf(cloLocal.lblOwnerIDValue.Text.Trim = String.Empty, 0, CType(cloLocal.lblOwnerIDValue.Text.Trim, Integer))
                    entityType = UIUtilsGen.EntityTypes.Owner
                    flagBarometerOwnerID = entityID
                    If cloLocal.lblFacilityIDValue.Text.Trim <> String.Empty Then
                        entityID = CType(cloLocal.lblFacilityIDValue.Text.Trim, Integer)
                        entityType = UIUtilsGen.EntityTypes.Facility
                        If cloLocal.nCurrentEventID > 0 Then
                            eventID = cloLocal.nCurrentEventID
                            eventType = UIUtilsGen.EntityTypes.ClosureEvent
                        End If
                    End If
                ElseIf strModuleTitle.ToUpper = "FINANCIAL" Then
                    Me.lblOwner.Visible = True
                    Me.lblOwnerInfo.Visible = True
                    Me.lblFacility.Visible = True
                    Me.lblFacilityID.Visible = True
                    Me.lblFacilityInfo.Visible = True
                    Dim finLocal As Financial = CType(ActiveMdiChild, Financial)
                    entityID = IIf(finLocal.lblOwnerIDValue.Text = String.Empty, 0, CType(finLocal.lblOwnerIDValue.Text, Integer))
                    entityType = UIUtilsGen.EntityTypes.Owner
                    flagBarometerOwnerID = entityID
                    If finLocal.lblFacilityIDValue.Text.Trim <> String.Empty Then
                        entityID = CType(finLocal.lblFacilityIDValue.Text.Trim, Integer)
                        entityType = UIUtilsGen.EntityTypes.Facility
                        If finLocal.nCurrentEventID > 0 Or finLocal.nCurrentEventID < -100 Then
                            eventID = finLocal.nCurrentEventID
                            eventType = UIUtilsGen.EntityTypes.FinancialEvent
                        End If
                    End If
                ElseIf strModuleTitle.ToUpper = "FEES" Then
                    Me.lblOwner.Visible = True
                    Me.lblOwnerInfo.Visible = True
                    Me.lblFacility.Visible = True
                    Me.lblFacilityID.Visible = True
                    Me.lblFacilityInfo.Visible = True
                    Dim feesLocal As Fees = CType(ActiveMdiChild, Fees)
                    entityID = IIf(feesLocal.lblOwnerIDValue.Text = String.Empty, 0, CType(feesLocal.lblOwnerIDValue.Text, Integer))
                    entityType = UIUtilsGen.EntityTypes.Owner
                    flagBarometerOwnerID = entityID
                    If feesLocal.lblFacilityIDValue.Text.Trim <> String.Empty Then
                        entityID = CType(feesLocal.lblFacilityIDValue.Text.Trim, Integer)
                        entityType = UIUtilsGen.EntityTypes.Facility
                    End If
                ElseIf strModuleTitle.ToUpper = "C & E" Then
                    Me.lblOwner.Visible = True
                    Me.lnklblPrevForm.Visible = True
                    Me.lblOwnerInfo.Visible = True
                    Me.lblFacility.Visible = True
                    Me.lnkLblNextForm.Visible = True
                    Me.lblFacilityID.Visible = True
                    Me.lblFacilityInfo.Visible = True
                    Dim caeLocal As CandE = CType(ActiveMdiChild, CandE)
                    entityID = IIf(caeLocal.lblOwnerIDValue.Text.Trim = String.Empty, 0, CType(caeLocal.lblOwnerIDValue.Text.Trim, Integer))
                    entityType = UIUtilsGen.EntityTypes.Owner
                    flagBarometerOwnerID = entityID
                    If caeLocal.lblFacilityIDValue.Text.Trim <> String.Empty Then
                        entityID = CType(caeLocal.lblFacilityIDValue.Text.Trim, Integer)
                        entityType = UIUtilsGen.EntityTypes.Facility
                    End If
                ElseIf strModuleTitle.ToUpper = "C&&E ADMIN" Then
                    'ElseIf strModuleTitle.ToUpper = "C " + "&" + " E ADMIN" Then
                    Me.lblOwner.Visible = False
                    Me.lnklblPrevForm.Visible = False
                    Me.lblOwnerInfo.Visible = False
                    Me.lblFacility.Visible = False
                    Me.lnkLblNextForm.Visible = False
                    Me.lblFacilityID.Visible = False
                    Me.lblFacilityInfo.Visible = False
                    entityID = 0
                    entityType = 0
                    flagBarometerOwnerID = 0
                ElseIf strModuleTitle.ToUpper = "COMPANY" Then
                    Me.lblOwner.Visible = True
                    Me.lblOwner.Text = "Company "
                    Me.lblOwnerInfo.Visible = True
                    Dim comLocal As Company = CType(ActiveMdiChild, Company)
                    entityID = comLocal.nCompanyID
                    entityType = UIUtilsGen.EntityTypes.Company
                    flagBarometerOwnerID = entityID
                ElseIf strModuleTitle.ToUpper = "CAP" Then
                    Me.lblOwner.Visible = False
                    Me.lnklblPrevForm.Visible = False
                    Me.lblOwnerInfo.Visible = False
                    Me.lblFacility.Visible = False
                    Me.lnkLblNextForm.Visible = False
                    Me.lblFacilityID.Visible = False
                    Me.lblFacilityInfo.Visible = False
                    entityID = 0
                    entityType = 0
                    flagBarometerOwnerID = 0
                End If
                LoadBarometer(entityID, entityType, IIf(strModuleTitle = "LUST", "Technical", strModuleTitle), ActiveMdiChild.Text, eventID, eventType)
            Else
                Me.pnlCommonReferenceArea.Visible = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ClearComboBox(ByVal frm As Form)
        Dim ActiveFrm As Object
        If TypeOf frm Is Technical Then
            ActiveFrm = CType(frm, MUSTER.Technical)
        ElseIf TypeOf frm Is Registration Then
            ActiveFrm = CType(frm, MUSTER.Registration)
        ElseIf TypeOf frm Is Closure Then
            ActiveFrm = CType(frm, Closure)
        ElseIf TypeOf frm Is Financial Then
            ActiveFrm = CType(frm, Financial)
        ElseIf TypeOf frm Is Fees Then
            ActiveFrm = CType(frm, Fees)
        ElseIf TypeOf frm Is CandE Then
            ActiveFrm = CType(frm, CandE)
        End If

        If pOwn.Facilities.FacilityType = 0 Then
            ActiveFrm.cmbFacilityType.SelectedIndex = -1
            ActiveFrm.cmbFacilityType.SelectedIndex = -1
        End If
        If Not TypeOf frm Is Fees Then
            If pOwn.Facilities.Datum = 0 Then
                ActiveFrm.cmbFacilityDatum.SelectedIndex = -1
                ActiveFrm.cmbFacilityDatum.SelectedIndex = -1
            End If
            If pOwn.Facilities.LocationType = 0 Then
                ActiveFrm.cmbFacilityLocationType.SelectedIndex = -1
                ActiveFrm.cmbFacilityLocationType.SelectedIndex = -1
            End If
            If pOwn.Facilities.Method = 0 Then
                ActiveFrm.cmbFacilityMethod.SelectedIndex = -1
                ActiveFrm.cmbFacilityMethod.SelectedIndex = -1
            End If
        End If
    End Sub
    Private Sub FillRegForm(ByVal nOwnerID As Integer, ByVal nFacilityId As Integer, ByVal Search_Type As String, Optional ByRef RegForm As Registration = Nothing, Optional ByRef TechForm As Technical = Nothing, Optional ByRef ClosrForm As Closure = Nothing, Optional ByRef FinancialForm As Financial = Nothing, Optional ByVal FeesForm As Fees = Nothing, Optional ByVal InspectionForm As Inspection = Nothing, Optional ByVal CandEForm As CandE = Nothing, Optional ByVal CandEMgmt As CandEManagement = Nothing, Optional ByVal CandEInspectors As Inspectors = Nothing, Optional ByVal CAPPreMonthlyForm As CAPPreMonthly = Nothing)
        Try
            Dim strModule As String
            If Not RegForm Is Nothing Then
                strModule = "REGISTRATION"
                If (Search_Type.IndexOf("Facility") > -1) Or (Search_Type.IndexOf("Ensite AIID") > -1) Then
                    RegForm.tbCntrlRegistration.SelectedTab = RegForm.tbPageFacilityDetail
                    RegForm.btnDeleteFacility.Enabled = True
                    RegForm.PopulateFacility(nFacilityId)
                Else
                    RegForm.tbCntrlRegistration.SelectedTab = RegForm.tbPageOwnerDetail
                    RegForm.Text = "Registration - Owner Detail (" & RegForm.txtOwnerName.Text & ")"
                    RegForm.PopulateOwner(nOwnerID, True)
                End If
                RegForm.MdiParent = Me
                RegForm.Show()
                RegForm.BringToFront()
                'RegForm.strOwnerWindowMode = RegForm.WindowMode.ForEdit
                HideShowOwnerInfo(True)
            ElseIf Not ClosrForm Is Nothing Then
                strModule = "CLOSURE"
                If Search_Type.IndexOf("Facility") > -1 Then
                    ClosrForm.tbCntrlClosure.SelectedTab = ClosrForm.tbPageFacilityDetail
                    ClosrForm.PopulateFacilityInfo(Integer.Parse(nFacilityId))
                Else
                    ClosrForm.tbCntrlClosure.SelectedTab = ClosrForm.tbPageOwnerDetail
                    ClosrForm.Text = "Closure - Owner Detail (" & ClosrForm.txtOwnerName.Text & ")"
                    ClosrForm.PopulateOwnerInfo(Integer.Parse(nOwnerID))
                End If
                ClosrForm.MdiParent = Me
                ClosrForm.Show()
                ClosrForm.BringToFront()
                ClosrForm.Tag = "Closure"
            ElseIf Not FinancialForm Is Nothing Then
                strModule = "FINANCIAL"
                If Search_Type.IndexOf("Facility") > -1 Then
                    FinancialForm.tbCntrlFinancial.SelectedTab = FinancialForm.tbPageFacilityDetail
                    FinancialForm.PopulateFacilityInfo(Integer.Parse(nFacilityId))
                Else
                    FinancialForm.tbCntrlFinancial.SelectedTab = FinancialForm.tbPageOwnerDetail
                    FinancialForm.Text = "Financial - Owner Detail (" & FinancialForm.txtOwnerName.Text & ")"
                    FinancialForm.PopulateOwnerInfo(Integer.Parse(nOwnerID))
                End If
                FinancialForm.MdiParent = Me
                FinancialForm.Show()
                FinancialForm.BringToFront()
                FinancialForm.Tag = "Financial"
            ElseIf Not FeesForm Is Nothing Then
                strModule = "FEES"
                If Search_Type.IndexOf("Facility") > -1 Then
                    FeesForm.tbCntrlFees.SelectedTab = FeesForm.tbPageFacilityDetail
                    FeesForm.PopulateFacilityInfo(Integer.Parse(nFacilityId))
                Else
                    FeesForm.tbCntrlFees.SelectedTab = FeesForm.tbPageOwnerDetail
                    FeesForm.Text = "Fees - Owner Detail (" & FeesForm.txtOwnerName.Text & ")"
                    FeesForm.PopulateOwnerInfo(Integer.Parse(nOwnerID))
                End If
                FeesForm.MdiParent = Me
                FeesForm.Show()
                FeesForm.BringToFront()
                FeesForm.Tag = "Fees"
            ElseIf Not InspectionForm Is Nothing Then
                strModule = "INSPECTION"
                InspectionForm.Populate(AppUser.UserKey, nOwnerID, nFacilityId, True)
                InspectionForm.MdiParent = Me
                InspectionForm.Show()
                InspectionForm.BringToFront()
                InspectionForm.Tag = "Inspection"
                HideShowOwnerInfo(False)
            ElseIf Not CandEForm Is Nothing Then
                strModule = "CAE"
                If Search_Type.IndexOf("Facility") > -1 Then
                    CandEForm.tbCntrlCandE.SelectedTab = CandEForm.tbPageFacilityDetail
                    CandEForm.PopulateFacility(Integer.Parse(nFacilityId))
                Else
                    CandEForm.tbCntrlCandE.SelectedTab = CandEForm.tbPageOwnerDetail
                    CandEForm.Text = "C & E - Owner Detail (" & CandEForm.txtOwnerName.Text & ")"
                    CandEForm.PopulateOwner(Integer.Parse(nOwnerID))
                End If
                CandEForm.MdiParent = Me
                CandEForm.Show()
                CandEForm.BringToFront()
                CandEForm.Tag = "C " + "&" + " E"
            ElseIf Not CandEMgmt Is Nothing Then
                strModule = "CAEMANAGEMENT"
                CandEMgmt.MdiParent = Me
                CandEMgmt.Show()
                CandEMgmt.BringToFront()
                CandEMgmt.Tag = "C and E Management"
            ElseIf Not CandEInspectors Is Nothing Then
                strModule = "CAEINSPECTORS"
                CandEInspectors.MdiParent = Me
                CandEInspectors.Show()
                CandEInspectors.BringToFront()
                CandEInspectors.Tag = "C and E Inspectors"
            ElseIf Not CAPPreMonthlyForm Is Nothing Then
                strModule = "CAP"
                CAPPreMonthlyForm.MdiParent = Me
                CAPPreMonthlyForm.Show()
                CAPPreMonthlyForm.BringToFront()
                CAPPreMonthlyForm.Tag = "CAP PreMonthly"
            Else
                strModule = "TECHNICAL"
                'Technical module.. 
                If Search_Type.IndexOf("Facility") > -1 Then
                    TechForm.tbCntrlTechnical.SelectedTab = TechForm.tbPageFacilityDetail
                    TechForm.PopFacility(Integer.Parse(nFacilityId))
                Else
                    TechForm.tbCntrlTechnical.SelectedTab = TechForm.tbPageOwnerDetail
                    TechForm.Text = "Technical - Owner Detail (" & TechForm.txtOwnerName.Text & ")"
                    TechForm.PopulateOwnerInfo(Integer.Parse(nOwnerID))
                End If
                TechForm.MdiParent = Me
                TechForm.Show()
                TechForm.BringToFront()
                TechForm.Tag = "Technical"
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

#Region "Flags"
    Private Sub SF_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles SF.FlagAdded
        If strModuleTitle = String.Empty Then
            LoadBarometer(entityID, entityType, [Module], ParentFormText)
        Else
            LoadBarometer(entityID, entityType, strModuleTitle, ParentFormText)
        End If
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        RefreshCalendarInfo()
        LoadDueToMeCalendar()
        LoadToDoCalendar()
    End Sub
#End Region

#Region "Sync Events"
    ' Recursively get all the subdirectories from the given 
    Private Function GetRecursiveFiles(ByVal sourceDir As String, ByRef arrDirectories As ArrayList)
        Dim sDir As String
        Dim sDirInfo As IO.DirectoryInfo

        Dim destDir As String
        Dim fRecursive As Boolean
        Dim overWrite As Boolean
        ' Add trailing separators to the supplied paths if they don't exist. 
        If Not sourceDir.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
            sourceDir &= System.IO.Path.DirectorySeparatorChar
        End If

        ' Get a list of directories from the current parent. 
        For Each sDir In System.IO.Directory.GetDirectories(sourceDir)
            sDirInfo = New System.IO.DirectoryInfo(sDir)
            GetFileNamesFromDirectory(sDirInfo.FullName, arrDirectories)
            ' Since we are in recursive mode, copy the children also 
            GetRecursiveFiles(sDirInfo.FullName, arrDirectories)
            sDirInfo = Nothing
        Next

    End Function
    ' Recursively get all files from the given directory
    Private Function GetFileNamesFromDirectory(ByVal sourceDir As String, ByRef arrFiles As ArrayList)
        Dim sFile As String
        Dim sFileInfo As IO.FileInfo

        ' Get the files from the current parent. 
        For Each sFile In System.IO.Directory.GetFiles(sourceDir)
            sFileInfo = New System.IO.FileInfo(sFile)
            ' only add files which begin with numbers

            If IsNumeric(sFileInfo.Name.Substring(0, 1)) And sFileInfo.Extension <> String.Empty Then
                arrFiles.Add(sFileInfo)
            End If
            sFileInfo = Nothing
        Next
    End Function
    Private Function GetModuleIDForName(ByVal ModuleName As String) As Integer
        Dim moduleID As Integer = 0
        Select Case ModuleName.ToUpper
            Case "REGISTRATION"
                moduleID = UIUtilsGen.ModuleID.Registration
            Case "C & E"
                moduleID = UIUtilsGen.ModuleID.CAE
            Case "TECHNICAL"
                moduleID = UIUtilsGen.ModuleID.Technical
            Case "INSPECTION"
                moduleID = UIUtilsGen.ModuleID.Inspection
            Case "FINANCIAL"
                moduleID = UIUtilsGen.ModuleID.Financial
            Case "CLOSURE"
                moduleID = UIUtilsGen.ModuleID.Closure
            Case "FEES"
                moduleID = UIUtilsGen.ModuleID.Fees
            Case "COMPANY"
                moduleID = UIUtilsGen.ModuleID.Company
            Case "CONTACTMANAGEMENT"
                moduleID = UIUtilsGen.ModuleID.ContactManagement
            Case "ADMIN"
                moduleID = UIUtilsGen.ModuleID.Admin
            Case "GLOBAL"
                moduleID = UIUtilsGen.ModuleID.[Global]
            Case "FEEADMIN"
                moduleID = UIUtilsGen.ModuleID.FeeAdmin
            Case "CAP PROCESS"
                moduleID = UIUtilsGen.ModuleID.CAPProcess
        End Select
        Return moduleID
    End Function
#End Region




    Private Sub MusterContainer_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If HoldClosing Then

            HoldClosing = False
            e.Cancel = True
        End If

        If Not TicklerTimerThread Is Nothing Then
            TicklerTimerThread.Stop()
            TicklerTimerThread.Dispose()
            TicklerTimerThread = Nothing
        End If

        If Not TicklerAutoOpenThread Is Nothing Then
            TicklerAutoOpenThread.Stop()
            TicklerAutoOpenThread.Dispose()
            TicklerAutoOpenThread = Nothing
        End If

    End Sub


    Private Sub BtnTickler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTickler.Click

        If Not TicklerScreen Is Nothing AndAlso Not TicklerScreen.Visible Then
            BeginInvoke(TicklerScreen.InvokeRefreshScreen)
        End If

    End Sub

    Private Sub ticklerAlertRequest(ByVal forced As Boolean) Handles TicklerScreen.DisplayAlert

        TAlert.StartAlert(forced)

    End Sub


    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Me.btnCloseAll.PerformClick()
    End Sub

End Class
