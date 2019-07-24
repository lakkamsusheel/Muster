Imports System.Data.SqlClient
Imports System.Text
Public Class Registration
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.Registration
    '   Provides the Registration functionality and UI for the application
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
    '  1.7        JVC2    02/08/05   Integrated with Padmaja's work 
    '  1.8        JVC2    02/13/05   Removed all extraneous comments and dead code
    '  1.9        JVC2    08/09/05   Added btnAddTank2.Focus to end of SetupModifyTankForm - issue 237
    '  1.91       TMF     02/18/2009   Remarked Line 7749 to allow registration to remember facility id
    '-------------------------------------------------------------------------------
    Inherits System.Windows.Forms.Form

#Region "User Defined Delegates"
    ' to handle datetimepicker validation
    Private Delegate Sub HandlerDelegate()
#End Region

#Region " User Defined Variables"
    Public MyGuid As New System.Guid
    Private bolLoading As Boolean = False
    Public nOwnerID, nFacilityID, nTankID, nCompartmentNumber, nPipeID As Integer
    Public bolAddOwner, bolAddFacility, bolAddTank, bolAddComp, bolAddPipe As Boolean
    Public mContainer As MusterContainer
    Private WithEvents AddressForm As Address
    Public strFacilityIdTags As String
    Private ugTankRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private ugPipeRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private dtNullDate As Date = CDate("01/01/0001")
    Private dtTodayPlus90Days As Date = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 4, CDate(Today.Month.ToString + "/1/" + Today.Year.ToString)))
    Private strFromCompanySearch As String = String.Empty
    Private vListSubstance, vListCerclaNumber As Infragistics.Win.ValueList
    Private dsContacts As DataSet
    Private strFilterString As String = String.Empty
    Public bolNewPersona As Boolean = False

    Private dTableTankCompartments As DataTable
    Private rp As New Remove_Pencil

    Private strXLContactName As String = String.Empty
    Private strXHContactName As String = String.Empty

    Public WithEvents pOwn As MUSTER.BusinessLogic.pOwner
    Public WithEvents pTank As MUSTER.BusinessLogic.pTank
    Public WithEvents pPipe As MUSTER.BusinessLogic.pPipe
    Private pCompany As MUSTER.BusinessLogic.pCompany
    Private pLicensee As MUSTER.BusinessLogic.pLicensee
    Private ttCERCLA As New ToolTip
    Private oReg As MUSTER.BusinessLogic.pRegistration
    Private oRegActinfo As MUSTER.Info.RegistrationActivityInfo
    Private pConStruct As MUSTER.BusinessLogic.pContactStruct
    Private autoChange As Boolean

    Private WithEvents SF As ShowFlags
    Private WithEvents oCompanySearch As CompanySearch
    Private WithEvents ContactFrm As Contacts
    Private WithEvents objCntSearch As ContactSearch
    Dim returnVal As String = String.Empty
#End Region

#Region " User Defined Events"
    'Public Event FlagsChanged(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String)
    'Public Event EnableDisable(ByVal flag As Boolean)
#End Region

#Region " Public Property"
    Public ReadOnly Property MC() As MusterContainer
        Get
            mContainer = Me.MdiParent
            If mContainer Is Nothing Then
                mContainer = New MusterContainer
            End If
            Return mContainer
        End Get
    End Property
    Public Property FormLoading() As Boolean
        Get
            Return bolLoading
        End Get
        Set(ByVal Value As Boolean)
            bolLoading = Value
        End Set
    End Property
    'Public Property OwnerWindowMode() As Integer
    '    Get
    '        Return strOwnerWindowMode
    '    End Get
    '    Set(ByVal Value As Integer)
    '        strOwnerWindowMode = Value
    '    End Set
    'End Property
    'Public Property TankCompartmentFormMode() As Int16
    '    Get
    '        Return nTankCompartmentFormMode
    '    End Get
    '    Set(ByVal Value As Int16)
    '        nTankCompartmentFormMode = Value
    '    End Set
    'End Property
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef oOwner As MUSTER.BusinessLogic.pOwner = Nothing, Optional ByVal ownerID As Int64 = 0, Optional ByVal facID As Int64 = 0)
        MyBase.New()

        bolLoading = True
        MyGuid = System.Guid.NewGuid

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionInstalled, dtNullDate)
        UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionLastTested, dtNullDate)
        UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionLastInspected, dtNullDate)
        UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionInstalled, dtNullDate)
        UIUtilsGen.SetDatePickerValue(dtPickSecondaryContainmentLastInspected, dtNullDate)
        UIUtilsGen.SetDatePickerValue(dtPickElectronicDeviceInspected, dtNullDate)
        UIUtilsGen.SetDatePickerValue(dtPickATGLastInspected, dtNullDate)


        UIUtilsGen.SetDatePickerValue(dtSheerValueTest, dtNullDate)
        UIUtilsGen.SetDatePickerValue(dtSecondaryContainmentInspected, dtNullDate)
        UIUtilsGen.SetDatePickerValue(dtElectronicDeviceInspected, dtNullDate)

        'Add any initialization after the InitializeComponent() call
        Cursor.Current = Cursors.AppStarting
        Try
            bolLoading = True
            If oOwner Is Nothing Then
                pOwn = New MUSTER.BusinessLogic.pOwner
            Else
                pOwn = oOwner
            End If
            pTank = New MUSTER.BusinessLogic.pTank
            pPipe = New MUSTER.BusinessLogic.pPipe
            pCompany = New MUSTER.BusinessLogic.pCompany
            pLicensee = New MUSTER.BusinessLogic.pLicensee
            oReg = New MUSTER.BusinessLogic.pRegistration
            pConStruct = New MUSTER.BusinessLogic.pContactStruct

            '
            'Need to tell the AppUser that we've instantiated another Registration form...
            '
            MusterContainer.AppUser.LogEntry("Registration", MyGuid.ToString)
            '
            ' The following line enables all forms to detect the visible form in the MDI container
            '
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Registration")

            UIUtilsGen.PopulateOwnerType(Me.cmbOwnerType, pOwn)
            UIUtilsGen.PopulateOrgEntityType(Me.cmbOwnerOrgEntityCode, pOwn)
            UIUtilsGen.PopulateFacilityType(Me.cmbFacilityType, pOwn.Facilities)
            UIUtilsGen.PopulateFacilityMethod(Me.cmbFacilityMethod, pOwn.Facilities)
            UIUtilsGen.PopulateFacilityDatum(Me.cmbFacilityDatum, pOwn.Facilities)
            UIUtilsGen.PopulateFacilityLocationType(Me.cmbFacilityLocationType, pOwn.Facilities)

            nOwnerID = ownerID
            nFacilityID = facID
            nTankID = 0
            nCompartmentNumber = 0
            nPipeID = 0

            bolAddOwner = False
            bolAddFacility = False
            bolAddTank = False
            bolAddComp = False
            bolAddPipe = False

            if nownerid > 0 then oReg.RetrieveByOwnerID(ownerID)

            ' The next three lines are necessary now since tank detail, pipe detail
            '   and the Second Tank/Pipe Grid on the Tank/Pipe Details tab occupy
            '   the same tab.  By doing so, the tabs and grid can be kept separate
            '   in the design environment.  Please do not alter their dock styles
            '   in the designer!!!
            '
            dgPipesAndTanks2.Dock = DockStyle.Fill
            tbCntrlPipe.Dock = DockStyle.Fill
            tbCntrlTank.Dock = DockStyle.Fill
     

        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = False
        End Try
        Cursor.Current = Cursors.Default
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        '
        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
        MusterContainer.AppUser.LogExit(MyGuid.ToString)
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
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents lblFinancial As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents IssueFin As System.Windows.Forms.ColumnHeader
    Friend WithEvents DateRecdFin As System.Windows.Forms.ColumnHeader
    Friend WithEvents DateDueFin As System.Windows.Forms.ColumnHeader
    Friend WithEvents DocumentFin As System.Windows.Forms.ColumnHeader
    Friend WithEvents CountDownFin As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lstViewFinancialIssue As System.Windows.Forms.ListView
    Friend WithEvents lblDocumentFin As System.Windows.Forms.Label
    Friend WithEvents CommentPMAct As System.Windows.Forms.ColumnHeader
    Friend WithEvents ActByDatePM As System.Windows.Forms.ColumnHeader
    Friend WithEvents FlagDatePM As System.Windows.Forms.ColumnHeader
    Friend WithEvents ActivitPM As System.Windows.Forms.ColumnHeader
    Friend WithEvents StartDatePM As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lstViewActivity As System.Windows.Forms.ListView
    Friend WithEvents lblActivityPM As System.Windows.Forms.Label
    Friend WithEvents CommentPM As System.Windows.Forms.ColumnHeader
    Friend WithEvents IssuePM As System.Windows.Forms.ColumnHeader
    Friend WithEvents DateRecdPM As System.Windows.Forms.ColumnHeader
    Friend WithEvents DateDue As System.Windows.Forms.ColumnHeader
    Friend WithEvents DocumentPM As System.Windows.Forms.ColumnHeader
    Friend WithEvents Countdown As System.Windows.Forms.ColumnHeader
    Friend WithEvents SelCol As System.Windows.Forms.ColumnHeader
    Friend WithEvents lstViewDocuments As System.Windows.Forms.ListView
    Friend WithEvents lblDocuments As System.Windows.Forms.Label
    Friend WithEvents lblProjectManagerIssues As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader25 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader24 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader23 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ListView6 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader34 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader35 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader36 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader37 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader38 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader47 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader26 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader27 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader39 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader40 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader41 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader42 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader43 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader44 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader45 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader46 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader69 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader70 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader71 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader72 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader73 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader74 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader75 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader76 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader77 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader78 As System.Windows.Forms.ColumnHeader
    Friend WithEvents tbPageSummary As System.Windows.Forms.TabPage
    Friend WithEvents Panel12 As System.Windows.Forms.Panel
    Friend WithEvents tbCntrlRegistration As System.Windows.Forms.TabControl
    Friend WithEvents pnl_FacilityDetail As System.Windows.Forms.Panel

    Private ListViewItem1 As System.Windows.Forms.ListViewItem
    Private ListViewItem2 As System.Windows.Forms.ListViewItem
    Private ListViewItem3 As System.Windows.Forms.ListViewItem
    Private ListViewItem4 As System.Windows.Forms.ListViewItem
    Private ListViewItem5 As System.Windows.Forms.ListViewItem
    Private ListViewItem6 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"T", "2", "In Use", "10,000", "Gasoline", "Bare Steel", "CP", "Galvonic", "GW/VP", "", "10/10/2001", "10/01/1999"}, -1)
    Private ListViewItem7 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"Edit", "Delete", "05/05/03", "", "This Site is under Inspection and will soon be determined"}, -1)
    Private ListViewItem8 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"Martha", "21 Hwy 80, Jackson", "(601) 924-9777", "mmartin@aol.com", "LUST", "Legal"}, -1)
    Private ListViewItem9 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"Jimmy Bob", "435 High St.", "(601) 978-0134", "JBob@aol.com", "Closure", "Signee"}, -1, System.Drawing.Color.Empty, System.Drawing.Color.Empty, New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
    Private ListViewItem10 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"1234", "Yellow Transport", "2515 Goodman Rd W", "Horn Lake", "De Soto", "3", "3", "3", "9"}, -1)
    Private ListViewItem11 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"2345", "Mapco Express", "25 Hwy 18", "Jackson", "Hinds", "4", "4", "4", "12"}, -1, System.Drawing.Color.Empty, System.Drawing.Color.Empty, New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
    Private ListViewItem12 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"Martha", "21 Hwy 80, Jackson", "(601) 924-9777", "mmartin@aol.com", "LUST", "Legal"}, -1)
    Private ListViewItem13 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"Jimmy Bob", "435 High St.", "(601) 978-0134", "JBob@aol.com", "Closure", "Signee"}, -1, System.Drawing.Color.Empty, System.Drawing.Color.Empty, New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
    Private ListViewItem14 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"05/05/03", "", "This Site is under Inspection and will soon be determined"}, -1)
    Private ListViewItem15 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"1015", "Yellow Transportation", "2515 Goodman Rd W", "Horn Lake", "02/03/02", "Donna"}, -1)
    Private ListViewItem16 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"3095", "Mapco Express", "25 Hwy 18", "Jackson", "04/18/02", "Donna"}, -1)
    Private ListViewItem17 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"05/11/03", "$500", "$500", "06/11/03", "$0", "", "Active"}, -1, System.Drawing.Color.Empty, System.Drawing.Color.Empty, New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
    Private ListViewItem18 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"03/11/02", "$500", "$500", "04/11/02", "$500", "03/23/02", "Inactive"}, -1)
    Private ListViewItem19 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"2003", "$300", "$150", "$0", "$0", "$300", "$150"}, -1)
    Private ListViewItem20 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"1001", "$10,000", "$9,000", "$1000", "Active"}, -1)
    Private ListViewItem21 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"1001", "11/12/03", "STFS", "Confirmed", "Lynn Chambers", "1", "Active", ""}, -1)
    Private ListViewItem22 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"1002", "11/07/02", "STFS", "Confirmed", "Lynn Chambers", "1", "Inactive", "11/12/02"}, -1, System.Drawing.Color.Empty, System.Drawing.Color.Empty, New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
    Private ListViewItem23 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"T", "1", "In Use", "10,000", "Gasoline", "Bare Steel", "CP", "Galvonic", "GW/VM", "", "10/10/2001", "10/01/2000"}, -1)
    Private ListViewItem24 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New System.Windows.Forms.ListViewItem.ListViewSubItem() {New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "P", System.Drawing.SystemColors.HotTrack, System.Drawing.SystemColors.ScrollBar, New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "ID", System.Drawing.SystemColors.HotTrack, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Status", System.Drawing.SystemColors.WindowText, System.Drawing.SystemColors.Window, New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Type"), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Substance"), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Material"), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Sec Option"), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "CP Type"), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "LD-Group1"), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "LD-Group2"), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Last Used"), New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Installed")}, -1)
    Private ListViewItem25 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"P", "2", "In Use", "Presurized", "Gasoline", "FRP", "None", "None", "GW/VP", "10/10/2001", "10/01/2000"}, -1, System.Drawing.Color.Empty, System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(192, Byte)), Nothing)
    Private ListViewItem26 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem("")
    Private ListViewItem27 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"T", "2", "In Use", "10,000", "Gasoline", "Bare Steel", "CP", "Galvonic", "GW/VP", "", "10/10/2001", "10/01/1999"}, -1)
    Private ListViewItem28 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"", "2", "In Use", "Presurized", "Gasoline", "FRP", "None", "None", "GW/VP", "10/10/2001", "10/01/2000"}, -1, System.Drawing.Color.Empty, System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(192, Byte)), Nothing)
    Private ListViewItem29 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"Martha", "21 Hwy 80, Jackson", "(601) 924-9777", "mmartin@aol.com", "LUST", "Legal"}, -1)
    Private ListViewItem30 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"Jimmy Bob", "435 High St.", "(601) 978-0134", "JBob@aol.com", "Closure", "Signee"}, -1, System.Drawing.Color.Empty, System.Drawing.Color.Empty, New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte)))
    Private ListViewItem31 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"05/05/03", "This Site is under Inspection and will soon be determined"}, -1)
    Friend WithEvents btnDeleteFacility As System.Windows.Forms.Button
    Friend WithEvents btnFacilitySave As System.Windows.Forms.Button
    Friend WithEvents btnDeleteOwner As System.Windows.Forms.Button
    Friend WithEvents pnlOwnerDetail As System.Windows.Forms.Panel
    Friend WithEvents btnSaveOwner As System.Windows.Forms.Button
    Friend WithEvents tbPageOwnerDetail As System.Windows.Forms.TabPage
    Friend WithEvents tbPageOwnerFacilities As System.Windows.Forms.TabPage
    Friend WithEvents tbCtrlOwner As System.Windows.Forms.TabControl
    Friend WithEvents ctxMenuTank As System.Windows.Forms.ContextMenu
    Friend WithEvents ctxMenuTankPipe As System.Windows.Forms.ContextMenu
    Friend WithEvents MI_EditTank As System.Windows.Forms.MenuItem
    Friend WithEvents MI_DeleteTank As System.Windows.Forms.MenuItem
    Friend WithEvents MI_CopyTank As System.Windows.Forms.MenuItem
    Friend WithEvents MI_EditTank1 As System.Windows.Forms.MenuItem
    Friend WithEvents MI_CopyTank1 As System.Windows.Forms.MenuItem
    Friend WithEvents MI_CopyPipe As System.Windows.Forms.MenuItem
    Friend WithEvents MI_DeleteTankPipe As System.Windows.Forms.MenuItem
    Friend WithEvents MI_DetachPipes As System.Windows.Forms.MenuItem
    Friend WithEvents MI_EditPipe As System.Windows.Forms.MenuItem
    Friend WithEvents LI_AttachPipes1 As System.Windows.Forms.MenuItem
    Friend WithEvents MI_AttachPipes As System.Windows.Forms.MenuItem
    Friend WithEvents LI_AddTankCompartment1 As System.Windows.Forms.MenuItem
    Friend WithEvents LI_AddTankCompartment As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MI_NewTank1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MI_NewTank As System.Windows.Forms.MenuItem
    Friend WithEvents tbPageOwnerContactList As System.Windows.Forms.TabPage
    Friend WithEvents ctxMenuTankCompartment As System.Windows.Forms.ContextMenu
    Friend WithEvents tbPageManageTank As System.Windows.Forms.TabPage
    Friend WithEvents pnlTankDetailMainDisplay As System.Windows.Forms.Panel
    Friend WithEvents lnkLblNextTank As System.Windows.Forms.LinkLabel
    Friend WithEvents lnkLblPrevTank As System.Windows.Forms.LinkLabel
    Friend WithEvents lblTankID As System.Windows.Forms.Label
    Friend WithEvents tbCntrlTank As System.Windows.Forms.TabControl
    Friend WithEvents tbPageTankDetail As System.Windows.Forms.TabPage
    Friend WithEvents pnlTankButtons As System.Windows.Forms.Panel
    Friend WithEvents btnTankSave As System.Windows.Forms.Button
    Friend WithEvents pnlTankDetail As System.Windows.Forms.Panel
    Friend WithEvents btnDeleteTank As System.Windows.Forms.Button
    Friend WithEvents btnCopyTankProfileToNew As System.Windows.Forms.Button
    Friend WithEvents lblOwnerID As System.Windows.Forms.Label
    Friend WithEvents lblOwnerPhone As System.Windows.Forms.Label
    Friend WithEvents lblOwnerAIID As System.Windows.Forms.Label
    Friend WithEvents btnOwnerCancel As System.Windows.Forms.Button
    Friend WithEvents lblOwnerCapParticipant As System.Windows.Forms.Label
    Friend WithEvents lblOwnerEmail As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerButtons As System.Windows.Forms.Panel
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePicker8 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker9 As System.Windows.Forms.DateTimePicker
    Friend WithEvents CheckBox10 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel16 As System.Windows.Forms.Panel
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePicker10 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker11 As System.Windows.Forms.DateTimePicker
    Friend WithEvents CheckBox11 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel17 As System.Windows.Forms.Panel
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents DateTimePicker12 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker13 As System.Windows.Forms.DateTimePicker
    Friend WithEvents CheckBox12 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel18 As System.Windows.Forms.Panel
    Friend WithEvents DateTimePicker14 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker15 As System.Windows.Forms.DateTimePicker
    Friend WithEvents CheckBox13 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel19 As System.Windows.Forms.Panel
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents pnlOwnerContactContainer As System.Windows.Forms.Panel
    Friend WithEvents tbPageTankPipe As System.Windows.Forms.TabPage
    Friend WithEvents pnlTankPipe As System.Windows.Forms.Panel
    Friend WithEvents tabCtrlFacilityTankPipe As System.Windows.Forms.TabControl
    Friend WithEvents myLabel As New Label
    'Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnTankCompCol8 As System.Windows.Forms.Button
    Friend WithEvents btnTankCompCol7 As System.Windows.Forms.Button
    Friend WithEvents btnTankCompCol6 As System.Windows.Forms.Button
    Friend WithEvents btnTankCompCol5 As System.Windows.Forms.Button
    Friend WithEvents btnTankCompCol4 As System.Windows.Forms.Button
    Friend WithEvents btnTankCompCol3 As System.Windows.Forms.Button
    Friend WithEvents btnTankCompCol2 As System.Windows.Forms.Button
    Friend WithEvents btnTankCompCol1 As System.Windows.Forms.Button
    Friend WithEvents pnlTankCompartmentHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlTankDescHead As System.Windows.Forms.Panel
    Friend WithEvents lblTankDescHead As System.Windows.Forms.Label
    Friend WithEvents lblTankDescDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlTankDescriptionTop As System.Windows.Forms.Panel
    Friend WithEvents cmbTankStatus As System.Windows.Forms.ComboBox
    Friend WithEvents lblTankStatus As System.Windows.Forms.Label
    Friend WithEvents pnlTankMaterialHead As System.Windows.Forms.Panel
    Friend WithEvents lblTankMaterialHead As System.Windows.Forms.Label
    Friend WithEvents lblTankMaterialDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlTankMaterial As System.Windows.Forms.Panel
    Friend WithEvents dtPickCPLastTested As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickCPInstalled As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickInteriorLiningInstalled As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickLastInteriorLinningInspection As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTankMaterial As System.Windows.Forms.Label
    Friend WithEvents lblTankOption As System.Windows.Forms.Label
    Friend WithEvents cmbTankMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTankOptions As System.Windows.Forms.ComboBox
    Friend WithEvents lblTankCPType As System.Windows.Forms.Label
    Friend WithEvents cmbTankCPType As System.Windows.Forms.ComboBox
    Friend WithEvents lblDateCPInstalled As System.Windows.Forms.Label
    Friend WithEvents lblDtTankLstTested As System.Windows.Forms.Label
    Friend WithEvents chkBxSpillProtected As System.Windows.Forms.CheckBox
    Friend WithEvents chkBxTightfillAdapters As System.Windows.Forms.CheckBox
    Friend WithEvents lblDateLnInteriorInstalled As System.Windows.Forms.Label
    Friend WithEvents lblDtLnInteriorLstInspect As System.Windows.Forms.Label
    Friend WithEvents pnlReleaseDetection As System.Windows.Forms.Panel
    Friend WithEvents lblTankReleaseHead As System.Windows.Forms.Label
    Friend WithEvents lblTankReleaseDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlTankRelease As System.Windows.Forms.Panel
    Friend WithEvents dtPickTankTightnessTest As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblRelseDetection As System.Windows.Forms.Label
    Friend WithEvents lblLTankTightnessTstDt As System.Windows.Forms.Label
    Friend WithEvents pnlInstallerOath As System.Windows.Forms.Panel
    Friend WithEvents lblTankInstallerOath As System.Windows.Forms.Label
    Friend WithEvents lblTankInstallerOathDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlTankInstallerOath As System.Windows.Forms.Panel
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents lblNoOfFacilities As System.Windows.Forms.Label
    Friend WithEvents pnllblDateofInstallation As System.Windows.Forms.Panel
    Friend WithEvents pnlTankInstallation As System.Windows.Forms.Panel
    Friend WithEvents lblTankInstallation As System.Windows.Forms.Label
    Friend WithEvents dtPickTankInstalled As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickPlannedInstallation As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTankInstallationPlanedFor As System.Windows.Forms.Label
    Friend WithEvents dtPickDatePlacedInService As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTankPlacedInServiceOn As System.Windows.Forms.Label
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents pnlTankTotalCapacity As System.Windows.Forms.Panel
    Friend WithEvents pnllblTankTotalCapacity As System.Windows.Forms.Panel
    Friend WithEvents lblTankTotalCapcityCaption As System.Windows.Forms.Label
    Friend WithEvents lblTankTotalCapacity As System.Windows.Forms.Label
    Friend WithEvents pnllblTankClosure As System.Windows.Forms.Panel
    Friend WithEvents pnlTankClosure As System.Windows.Forms.Panel
    Friend WithEvents lblTankClosureCaption As System.Windows.Forms.Label
    Friend WithEvents lblTankClosure As System.Windows.Forms.Label
    Friend WithEvents lblDateClosed As System.Windows.Forms.Label
    Friend WithEvents dtPickLastUsed As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDtLastUsed As System.Windows.Forms.Label
    Friend WithEvents lblClosuredate As System.Windows.Forms.Label
    Friend WithEvents lblDateClosureRecvd As System.Windows.Forms.Label
    Friend WithEvents lblDateTankClosureRecvValue As System.Windows.Forms.Label
    Friend WithEvents lblTankClosureStatus As System.Windows.Forms.Label
    Friend WithEvents lblTankClosureStatusValue As System.Windows.Forms.Label
    Friend WithEvents lblTankInertFill As System.Windows.Forms.Label
    Friend WithEvents lblTankInertFillValue As System.Windows.Forms.Label
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents pnlOwnerBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlMain As System.Windows.Forms.Panel
    Friend WithEvents pnlOwnerFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlFacilityTankPipeButton As System.Windows.Forms.Panel
    Friend WithEvents lblDateReceived As System.Windows.Forms.Label
    Friend WithEvents lblFacilitySigOnNF As System.Windows.Forms.Label
    Friend WithEvents lblFacilityFuelBrand As System.Windows.Forms.Label
    Friend WithEvents lblFacilityStatus As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongitude As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatitude As System.Windows.Forms.Label
    Friend WithEvents lblFacilityType As System.Windows.Forms.Label
    Friend WithEvents lblfacilityAIID As System.Windows.Forms.Label
    Friend WithEvents lblFacilityID As System.Windows.Forms.Label
    Friend WithEvents lblFacilityPhone As System.Windows.Forms.Label
    Friend WithEvents lblFacilityName As System.Windows.Forms.Label
    Friend WithEvents txtfacilityPhone As System.Windows.Forms.TextBox
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents lblOwnerType As System.Windows.Forms.Label
    Friend WithEvents lblDateofInstallation As System.Windows.Forms.Label
    Friend WithEvents chkDeliveriesLimited As System.Windows.Forms.CheckBox
    Friend WithEvents chkOverFilledProtected As System.Windows.Forms.CheckBox
    Friend WithEvents chkEmergencyPower As System.Windows.Forms.CheckBox
    Friend WithEvents txtFacilityFax As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityFax As System.Windows.Forms.Label
    Friend WithEvents lblOwnerStatus As System.Windows.Forms.Label
    Friend WithEvents lblTotalNoOfTanks As System.Windows.Forms.Label
    Friend WithEvents lblFacilitySIC As System.Windows.Forms.Label
    Friend WithEvents lblTankOverfillProtectionType As System.Windows.Forms.Label
    Friend WithEvents tbPageFacilityDetail As System.Windows.Forms.TabPage
    Friend WithEvents lblOwnerName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerAddress As System.Windows.Forms.Label
    Friend WithEvents lblOwnerOrgEntityCode As System.Windows.Forms.Label
    Friend WithEvents btnOwnerNameCancel As System.Windows.Forms.Button
    Friend WithEvents btnOwnerNameOK As System.Windows.Forms.Button
    Friend WithEvents lblOwnerOrgName As System.Windows.Forms.Label
    Friend WithEvents btnOwnerNameSearch As System.Windows.Forms.Button
    Friend WithEvents lblOwnerNameSuffix As System.Windows.Forms.Label
    Friend WithEvents lblOwnerNameTitle As System.Windows.Forms.Label
    Friend WithEvents lblOwnerMiddleName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerLastName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerFirstName As System.Windows.Forms.Label
    Friend WithEvents lblFacilityAddress As System.Windows.Forms.Label
    Friend WithEvents txtFacilityZip As System.Windows.Forms.TextBox
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Friend WithEvents lblTankType As System.Windows.Forms.Label
    Friend WithEvents cmbTankType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTankCompCercla As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTankCompSubstance As System.Windows.Forms.ComboBox
    Friend WithEvents lblTankCapacity As System.Windows.Forms.Label
    Friend WithEvents chkTankCompartment As System.Windows.Forms.CheckBox
    Friend WithEvents lblTankCompartmentNumber As System.Windows.Forms.Label
    Friend WithEvents mnuAddCompPipe As System.Windows.Forms.MenuItem
    Friend WithEvents mnuShowCompPipes As System.Windows.Forms.MenuItem
    Friend WithEvents cmbTankOverfillProtectionType As System.Windows.Forms.ComboBox
    Friend WithEvents dtPickTankInstallerSigned As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTankInstallerDtSigned As System.Windows.Forms.Label
    Friend WithEvents chkTankDrpTubeInvControl As System.Windows.Forms.CheckBox
    Friend WithEvents cmbTankReleaseDetection As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacilityLatDegree As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongDegree As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatSec As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatMin As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongMin As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongSec As System.Windows.Forms.Label
    Friend WithEvents lblFacilityDatum As System.Windows.Forms.Label
    Friend WithEvents lblFacilityMethod As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLocationType As System.Windows.Forms.Label
    Friend WithEvents lblPowerOff As System.Windows.Forms.Label
    Friend WithEvents lnkLblPrevFacility As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLblCAPSignup As System.Windows.Forms.LinkLabel
    Friend WithEvents dGridCompartments As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents txtTankFacilityID As System.Windows.Forms.TextBox
    Friend WithEvents btnAddTank As System.Windows.Forms.Button
    Friend WithEvents btnTransferOwnership As System.Windows.Forms.Button
    Friend WithEvents ll As System.Windows.Forms.Label
    Friend WithEvents btnAddPipe As System.Windows.Forms.Button
    Friend WithEvents cmbTankManufacturer As System.Windows.Forms.ComboBox
    Friend WithEvents lblTankManufacturer As System.Windows.Forms.Label
    Friend WithEvents cmbTankInertFill As System.Windows.Forms.ComboBox
    Friend WithEvents lblNonCompTankCapacity As System.Windows.Forms.Label
    Friend WithEvents txtNonCompTankCapacity As System.Windows.Forms.TextBox
    Friend WithEvents lblTankManifold As System.Windows.Forms.Label
    Friend WithEvents lblTankFuelType As System.Windows.Forms.Label
    Friend WithEvents lblTankCercla As System.Windows.Forms.Label
    Friend WithEvents lblTankSubstance As System.Windows.Forms.Label
    Friend WithEvents pnlNonCompProperties As System.Windows.Forms.Panel
    Friend WithEvents cmbTanksubstance As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTankFuelType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTankCercla As System.Windows.Forms.ComboBox
    Friend WithEvents lblTankManifoldValue As System.Windows.Forms.Label
    Friend WithEvents tpPreviouslyOwnedOwners As System.Windows.Forms.TabPage
    Friend WithEvents lblNoOfPreviouslyOwnedFacilitiesValue As System.Windows.Forms.Label
    Friend WithEvents lblNoOfPreviouslyOwnedFacilities As System.Windows.Forms.Label
    Friend WithEvents btnAddExistingPipe As System.Windows.Forms.Button
    Friend WithEvents btnRegister As System.Windows.Forms.Button
    Friend WithEvents lblCAPStatus As System.Windows.Forms.Label
    Friend WithEvents lblTankOriginalStatusID As System.Windows.Forms.Label
    Friend WithEvents btnOwnerFlag As System.Windows.Forms.Button
    Friend WithEvents btnOwnerComment As System.Windows.Forms.Button
    Friend WithEvents tbPrevFacs As System.Windows.Forms.TabPage
    Friend WithEvents pnlPrevOwnedFacsCount As System.Windows.Forms.Panel
    Friend WithEvents lblNoofPreviousFacilities As System.Windows.Forms.Label
    Friend WithEvents lblPreviousFacilitiesCountValue As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnExpand As System.Windows.Forms.Button
    Friend WithEvents btnFacComments As System.Windows.Forms.Button
    Friend WithEvents btnFacFlags As System.Windows.Forms.Button
    Friend WithEvents tbPagePipeDetail As System.Windows.Forms.TabPage
    Friend WithEvents pnlPipeButtons As System.Windows.Forms.Panel
    Friend WithEvents btnCopyPipeProfile As System.Windows.Forms.Button
    Friend WithEvents btnDeletePipe As System.Windows.Forms.Button
    Friend WithEvents btnPipeSave As System.Windows.Forms.Button
    Friend WithEvents pnlPipeDetail As System.Windows.Forms.Panel
    Friend WithEvents pnlPipeClosure As System.Windows.Forms.Panel
    Friend WithEvents cmbPipeInertFill As System.Windows.Forms.ComboBox
    Friend WithEvents dtPickPipeLastUsed As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipeLastUsed As System.Windows.Forms.Label
    Friend WithEvents lblPipeClosedOn As System.Windows.Forms.Label
    Friend WithEvents lblPipeClosedOnDate As System.Windows.Forms.Label
    Friend WithEvents lblPipeInertFillValue As System.Windows.Forms.Label
    Friend WithEvents lblPipeInertFill As System.Windows.Forms.Label
    Friend WithEvents lblPipeClosureStatusValue As System.Windows.Forms.Label
    Friend WithEvents lblPipeClosureRcvdDateValue As System.Windows.Forms.Label
    Friend WithEvents lblPipeClosureRcvd As System.Windows.Forms.Label
    Friend WithEvents Panel20 As System.Windows.Forms.Panel
    Friend WithEvents lblPipeClosureCaption As System.Windows.Forms.Label
    Friend WithEvents lblPipeClosure As System.Windows.Forms.Label
    Friend WithEvents pnlPipeInstallerOath As System.Windows.Forms.Panel
    Friend WithEvents dtPickPipeSigned As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDtPipeSigned As System.Windows.Forms.Label
    Friend WithEvents Panel9 As System.Windows.Forms.Panel
    Friend WithEvents lblPipeInstallerOath As System.Windows.Forms.Label
    Friend WithEvents lblPipeInstallerOathDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlPipeRelease As System.Windows.Forms.Panel
    Friend WithEvents grpBxReleaseDetectionGroup2 As System.Windows.Forms.GroupBox
    Friend WithEvents dtPickPipeLeakDetectorTest As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipeReleaseDetection2 As System.Windows.Forms.Label
    Friend WithEvents cmbPipeReleaseDetection2 As System.Windows.Forms.ComboBox
    Friend WithEvents lblPipeALLDTestDate As System.Windows.Forms.Label
    Friend WithEvents grpBxPipeReleaseDetectionGroup1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtPickPipeTightnessTest As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipeReleaseDetection As System.Windows.Forms.Label
    Friend WithEvents cmbPipeReleaseDetection1 As System.Windows.Forms.ComboBox
    Friend WithEvents lblLPipeTightnessTstDt As System.Windows.Forms.Label
    Friend WithEvents Panel11 As System.Windows.Forms.Panel
    Friend WithEvents lblPipeReleaseHead As System.Windows.Forms.Label
    Friend WithEvents lblPipeReleaseDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlPipeType As System.Windows.Forms.Panel
    Friend WithEvents cmbPipeType As System.Windows.Forms.ComboBox
    Friend WithEvents lblPipeCapacity As System.Windows.Forms.Label
    Friend WithEvents Panel15 As System.Windows.Forms.Panel
    Friend WithEvents lblPipeTypeCaption As System.Windows.Forms.Label
    Friend WithEvents lblPipeType As System.Windows.Forms.Label
    Friend WithEvents pnlPipeMaterial As System.Windows.Forms.Panel
    Friend WithEvents lblPipeManufacturer As System.Windows.Forms.Label
    Friend WithEvents cmbPipeManufacturerID As System.Windows.Forms.ComboBox
    Friend WithEvents dtPickPipeTerminationCPLastTested As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipeTerminationCPLastTest As System.Windows.Forms.Label
    Friend WithEvents dtPickPipeTerminationCPInstalled As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipeTerminationCPInstalled As System.Windows.Forms.Label
    Friend WithEvents grpBxPipeTerminationAtDispenser As System.Windows.Forms.GroupBox
    Friend WithEvents cmbPipeTerminationDispenserCPType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbPipeTerminationDispenserType As System.Windows.Forms.ComboBox
    Friend WithEvents lblPipeTerminationDispenserCPType As System.Windows.Forms.Label
    Friend WithEvents lblPipeTerminationDispenserType As System.Windows.Forms.Label
    Friend WithEvents grpBxPipeTerminationAtTank As System.Windows.Forms.GroupBox
    Friend WithEvents cmbPipeTerminationTankCPType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbPipeTerminationTankType As System.Windows.Forms.ComboBox
    Friend WithEvents lblPipeTerminationTankCPType As System.Windows.Forms.Label
    Friend WithEvents lblPipeTerminationTankType As System.Windows.Forms.Label
    Friend WithEvents grpBxPipeContainmentSumpsLocation As System.Windows.Forms.GroupBox
    Friend WithEvents chkPipeSumpsAtTank As System.Windows.Forms.CheckBox
    Friend WithEvents chkPipeSumpsAtDispenser As System.Windows.Forms.CheckBox
    Friend WithEvents dtPickPipeCPLastTest As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickPipeCPInstalled As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipeMaterial As System.Windows.Forms.Label
    Friend WithEvents lblPipeOption As System.Windows.Forms.Label
    Friend WithEvents cmbPipeMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents cmbPipeOptions As System.Windows.Forms.ComboBox
    Friend WithEvents lblPipeCPType As System.Windows.Forms.Label
    Friend WithEvents cmbPipeCPType As System.Windows.Forms.ComboBox
    Friend WithEvents lblDatePipeCPInstalled As System.Windows.Forms.Label
    Friend WithEvents lblDtPipeLastTested As System.Windows.Forms.Label
    Friend WithEvents pnlPipeMaterialHead As System.Windows.Forms.Panel
    Friend WithEvents lblPipeMaterialHead As System.Windows.Forms.Label
    Friend WithEvents lblPipeMaterialDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlPipeDateOfInstallation As System.Windows.Forms.Panel
    Friend WithEvents lblPipeFuelTypeValue As System.Windows.Forms.Label
    Friend WithEvents lblPipeCerclaNoValue As System.Windows.Forms.Label
    Friend WithEvents lblPipeCerclaNo As System.Windows.Forms.Label
    Friend WithEvents lblPipeSubstance As System.Windows.Forms.Label
    Friend WithEvents lblPipeSubstanceValue As System.Windows.Forms.Label
    Friend WithEvents dtPickDatePipePlacedInService As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipePlacedInServiceOn As System.Windows.Forms.Label
    Friend WithEvents lblPipeInstallationPlanedFor As System.Windows.Forms.Label
    Friend WithEvents dtPickPipePlannedInstallation As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipeInstalledOn As System.Windows.Forms.Label
    Friend WithEvents dtPickPipeInstalled As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPipeFuelType As System.Windows.Forms.Label
    Friend WithEvents Panel13 As System.Windows.Forms.Panel
    Friend WithEvents lblPipeDateOfInstallationCaption As System.Windows.Forms.Label
    Friend WithEvents lblPipeDateOfInstallation As System.Windows.Forms.Label
    Friend WithEvents pnlPipeDescription As System.Windows.Forms.Panel
    Friend WithEvents cmbPipeStatus As System.Windows.Forms.ComboBox
    Friend WithEvents lblPipeStatus As System.Windows.Forms.Label
    Friend WithEvents pnlPipeDescHead As System.Windows.Forms.Panel
    Friend WithEvents lblPipeDescHead As System.Windows.Forms.Label
    Friend WithEvents lblPipeDescDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlPipeDetailHeader As System.Windows.Forms.Panel
    Friend WithEvents lblPipeIndex As System.Windows.Forms.Label
    Friend WithEvents lblPipeCompartmentIndex As System.Windows.Forms.Label
    Friend WithEvents lblPipeCompartment As System.Windows.Forms.Label
    Friend WithEvents lblPipeID As System.Windows.Forms.Label
    Friend WithEvents lblPipeTankID As System.Windows.Forms.Label
    Friend WithEvents pnlTankDetailHeader As System.Windows.Forms.Panel
    Friend WithEvents dgPipesAndTanks2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents tbCntrlPipe As System.Windows.Forms.TabControl
    Friend WithEvents lblTankIDValue As System.Windows.Forms.Label
    Friend WithEvents lblTankIDValue2 As System.Windows.Forms.Label
    Friend WithEvents pnlTankCount2 As System.Windows.Forms.Panel
    Friend WithEvents lblTotalNoOfTanksValue2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblPipeIDValue As System.Windows.Forms.Label
    Friend WithEvents lblPipeTankIDValue As System.Windows.Forms.Label
    Friend WithEvents btnExpandTP2 As System.Windows.Forms.Button
    Friend WithEvents btnToTank As System.Windows.Forms.Button
    Friend WithEvents btnToPipe As System.Windows.Forms.Button
    Friend WithEvents lblTankCountVal As System.Windows.Forms.Label
    Friend WithEvents lblTankCountVal2 As System.Windows.Forms.Label
    Friend WithEvents btnAddTank2 As System.Windows.Forms.Button
    Friend WithEvents lnkLblNextFac As System.Windows.Forms.LinkLabel
    Friend WithEvents lblUpcomingInstallDate As System.Windows.Forms.Label
    Friend WithEvents lblTankInstalledOn As System.Windows.Forms.Label
    Friend WithEvents btnPipeComments As System.Windows.Forms.Button
    Friend WithEvents btnTankComments As System.Windows.Forms.Button
    Friend WithEvents btnOwnerNameClose As System.Windows.Forms.Button
    Friend WithEvents btnFacilityCancel As System.Windows.Forms.Button
    Friend WithEvents btnTankCancel As System.Windows.Forms.Button
    Friend WithEvents btnPipeCancel As System.Windows.Forms.Button
    Public WithEvents lblOwnerIDValue As System.Windows.Forms.Label
    Public WithEvents cmbFacilityType As System.Windows.Forms.ComboBox
    Public WithEvents cmbOwnerType As System.Windows.Forms.ComboBox
    Public WithEvents txtOwnerAIID As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerEmail As System.Windows.Forms.TextBox
    Public WithEvents lblNoOfFacilitiesValue As System.Windows.Forms.Label
    Public WithEvents dtPickFacilityRecvd As System.Windows.Forms.DateTimePicker
    Public WithEvents lblFacilityStatusValue As System.Windows.Forms.Label
    Public WithEvents lblFacilityIDValue As System.Windows.Forms.Label
    Public WithEvents txtFacilityAIID As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityName As System.Windows.Forms.TextBox
    Public WithEvents chkSignatureofNF As System.Windows.Forms.CheckBox
    Public WithEvents lblTotalNoOfTanksValue As System.Windows.Forms.Label
    Public WithEvents txtOwnerName As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerAddress As System.Windows.Forms.TextBox
    Public WithEvents cmbOwnerOrgEntityCode As System.Windows.Forms.ComboBox
    Public WithEvents txtOwnerOrgName As System.Windows.Forms.TextBox
    Public WithEvents rdOwnerOrg As System.Windows.Forms.RadioButton
    Public WithEvents rdOwnerPerson As System.Windows.Forms.RadioButton
    Public WithEvents cmbOwnerNameSuffix As System.Windows.Forms.ComboBox
    Public WithEvents cmbOwnerNameTitle As System.Windows.Forms.ComboBox
    Public WithEvents txtOwnerMiddleName As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerLastName As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerFirstName As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityAddress As System.Windows.Forms.TextBox
    Public WithEvents mskTxtOwnerPhone As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone2 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtFacilityFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtFacilityPhone As AxMSMask.AxMaskEdBox
    Public WithEvents txtFacilityLongDegree As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatDegree As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLongMin As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatMin As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatSec As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLongSec As System.Windows.Forms.TextBox
    Public WithEvents cmbFacilityDatum As System.Windows.Forms.ComboBox
    Public WithEvents cmbFacilityMethod As System.Windows.Forms.ComboBox
    Public WithEvents cmbFacilityLocationType As System.Windows.Forms.ComboBox
    Public WithEvents lblCAPParticipationLevel As System.Windows.Forms.Label
    Public WithEvents chkLUSTSite As System.Windows.Forms.CheckBox
    Public WithEvents dtFacilityPowerOff As System.Windows.Forms.DateTimePicker
    Public WithEvents lblOwnerActiveOrNot As System.Windows.Forms.Label
    Public WithEvents dgPipesAndTanks As Infragistics.Win.UltraWinGrid.UltraGrid
    Public WithEvents chkOwnerAgencyInterest As System.Windows.Forms.CheckBox
    Public WithEvents txtFuelBrand As System.Windows.Forms.TextBox
    Public WithEvents chkCAPCandidate As System.Windows.Forms.CheckBox
    Public WithEvents lblCAPStatusValue As System.Windows.Forms.Label
    Public WithEvents lblOwnerLastEditedOn As System.Windows.Forms.Label
    Public WithEvents lblOwnerLastEditedBy As System.Windows.Forms.Label
    Public WithEvents dtPickUpcomingInstallDateValue As System.Windows.Forms.DateTimePicker
    Public WithEvents chkUpcomingInstall As System.Windows.Forms.CheckBox
    Public WithEvents ugFacilityList As Infragistics.Win.UltraWinGrid.UltraGrid
    Public WithEvents chkCAPParticipant As System.Windows.Forms.CheckBox
    Public WithEvents lblNewOwnerSnippetValue As System.Windows.Forms.Label
    Public WithEvents lblDateTransfered As System.Windows.Forms.Label
    Public WithEvents pnlOwnerName As System.Windows.Forms.Panel
    Public WithEvents pnlOwnerNameButton As System.Windows.Forms.Panel
    Public WithEvents pnlOwnerOrg As System.Windows.Forms.Panel
    Public WithEvents pnlPersonOrganization As System.Windows.Forms.Panel
    Public WithEvents pnlOwnerPerson As System.Windows.Forms.Panel
    Public WithEvents txtDueByNF As System.Windows.Forms.TextBox

    '----- Friend controls --------------------------------------------------------------------------
    Friend WithEvents pnlOwnerContactButtons As System.Windows.Forms.Panel
    Friend WithEvents ugOwnerContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnOwnerModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAddSearchContact As System.Windows.Forms.Button
    Friend WithEvents tbPageFacilityContactList As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerContactHeader As System.Windows.Forms.Panel
    Friend WithEvents chkOwnerShowActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents lblOwnerContacts As System.Windows.Forms.Label
    Friend WithEvents pnlFacilityContactHeader As System.Windows.Forms.Panel
    Friend WithEvents chkFacilityShowActiveContactOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkFacilityShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkFacilityShowContactsforAllModule As System.Windows.Forms.CheckBox
    Friend WithEvents lblFacilityContacts As System.Windows.Forms.Label
    Friend WithEvents pnlFacilityContactContainer As System.Windows.Forms.Panel
    Friend WithEvents ugFacilityContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pnlFacilityContactBottom As System.Windows.Forms.Panel
    Friend WithEvents btnFacilityModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnFacilityDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnFacilityAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnFacilityAddSearchContact As System.Windows.Forms.Button
    Friend WithEvents txtTankCompany As System.Windows.Forms.TextBox
    Friend WithEvents lblLicenseeName As System.Windows.Forms.Label
    Friend WithEvents lblTankInstallerCompany As System.Windows.Forms.Label
    Friend WithEvents txtPipeCompanyName As System.Windows.Forms.TextBox
    Friend WithEvents lblPipeInstallerName As System.Windows.Forms.Label
    Friend WithEvents lblPipeInstallerCompanyName As System.Windows.Forms.Label
    Friend WithEvents txtTankCompartmentNumber As System.Windows.Forms.Label
    Friend WithEvents txtTankCapacity As System.Windows.Forms.Label
    Friend WithEvents cmbTankCerclaDesc As System.Windows.Forms.ComboBox
    Friend WithEvents chkEmergencyPowerPipe As System.Windows.Forms.CheckBox
    Friend WithEvents lblEmergencyPowerPipe As System.Windows.Forms.Label
    Friend WithEvents lblActiveLust As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblLicenseeSearch As System.Windows.Forms.Label
    Friend WithEvents txtLicensee As System.Windows.Forms.TextBox
    Friend WithEvents lblPipeLicenseeCompanySearch As System.Windows.Forms.Label
    Friend WithEvents txtPipeLicensee As System.Windows.Forms.TextBox
    Friend WithEvents cmbTankClosureType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbPipeClosureType As System.Windows.Forms.ComboBox
    Friend WithEvents lblPipeClosureStatus As System.Windows.Forms.Label
    Friend WithEvents ugPrevOwnedFacs As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents lblNoOfOwners As System.Windows.Forms.Label
    Friend WithEvents lblNoofOwnersValue As System.Windows.Forms.Label
    Friend WithEvents ugPreviouslyOwnedOwners As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblCERCLAtt As System.Windows.Forms.Label
    Friend WithEvents txtFacilityLicensee As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityLicenseeSearch As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLicensee As System.Windows.Forms.Label
    Friend WithEvents txtFacilityCompany As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityCompany As System.Windows.Forms.Label
    Friend WithEvents tbPageOwnerDocuments As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityDocuments As System.Windows.Forms.TabPage
    Friend WithEvents UCFacilityDocuments As MUSTER.DocumentViewControl
    Friend WithEvents UCOwnerDocuments As MUSTER.DocumentViewControl
    Friend WithEvents pnlOwnerSummaryDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlOwnerSummaryHeader As System.Windows.Forms.Panel
    Public WithEvents UCOwnerSummary As MUSTER.OwnerSummary
    Friend WithEvents btnEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnLabels As System.Windows.Forms.Button
    Friend WithEvents btnFacilityLabels As System.Windows.Forms.Button
    Friend WithEvents btnFacilityEnvelopes As System.Windows.Forms.Button
    Public WithEvents txtFacilitySIC As System.Windows.Forms.Label
    Friend WithEvents chkBoxReplacementTank As System.Windows.Forms.CheckBox
    Public WithEvents txtFacilityNameForEnsite As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityNameForEnsite As System.Windows.Forms.Label
    Friend WithEvents LinkLblCAPSignupFac As System.Windows.Forms.LinkLabel
    Friend WithEvents btnDetachPipes As System.Windows.Forms.Button
    Friend WithEvents chkProhibition As System.Windows.Forms.CheckBox
    Friend WithEvents chkTankProhibition As System.Windows.Forms.CheckBox
    Friend WithEvents LblDateSpillPreventionInstalled As System.Windows.Forms.Label
    Friend WithEvents LblDateSpillPreventionLastTested As System.Windows.Forms.Label
    Friend WithEvents dtPickSpillPreventionInstalled As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblDateOverfillPreventionLastInspected As System.Windows.Forms.Label
    Friend WithEvents lbLDateSecondaryContainmentLastInspected As System.Windows.Forms.Label
    Friend WithEvents dtElectronicDeviceInspected As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblElectronicDeviceInspected As System.Windows.Forms.Label
    Friend WithEvents LblDateElectronicDeviceInspected As System.Windows.Forms.Label
    Friend WithEvents LblDateATGLastInspected As System.Windows.Forms.Label
    Friend WithEvents dtPickSpillPreventionLastTested As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickOverfillPreventionLastInspected As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickSecondaryContainmentLastInspected As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickElectronicDeviceInspected As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickATGLastInspected As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickOverfillPreventionInstalled As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblDateOverfillPreventionInstalled As System.Windows.Forms.Label
    Friend WithEvents dtSecondaryContainmentInspected As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblSecondaryContainmentInspected As System.Windows.Forms.Label
    Friend WithEvents dtSheerValueTest As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblSheerValueTestDate As System.Windows.Forms.Label
    Public WithEvents dtPickAssess As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblAssessDate As System.Windows.Forms.Label


    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Registration))
        Me.Panel10 = New System.Windows.Forms.Panel
        Me.lblFinancial = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.IssueFin = New System.Windows.Forms.ColumnHeader
        Me.DateRecdFin = New System.Windows.Forms.ColumnHeader
        Me.DateDueFin = New System.Windows.Forms.ColumnHeader
        Me.DocumentFin = New System.Windows.Forms.ColumnHeader
        Me.CountDownFin = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.lstViewFinancialIssue = New System.Windows.Forms.ListView
        Me.lblDocumentFin = New System.Windows.Forms.Label
        Me.ActByDatePM = New System.Windows.Forms.ColumnHeader
        Me.FlagDatePM = New System.Windows.Forms.ColumnHeader
        Me.ActivitPM = New System.Windows.Forms.ColumnHeader
        Me.StartDatePM = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.lstViewActivity = New System.Windows.Forms.ListView
        Me.lblActivityPM = New System.Windows.Forms.Label
        Me.CommentPM = New System.Windows.Forms.ColumnHeader
        Me.IssuePM = New System.Windows.Forms.ColumnHeader
        Me.DateRecdPM = New System.Windows.Forms.ColumnHeader
        Me.DateDue = New System.Windows.Forms.ColumnHeader
        Me.DocumentPM = New System.Windows.Forms.ColumnHeader
        Me.Countdown = New System.Windows.Forms.ColumnHeader
        Me.SelCol = New System.Windows.Forms.ColumnHeader
        Me.lstViewDocuments = New System.Windows.Forms.ListView
        Me.lblDocuments = New System.Windows.Forms.Label
        Me.lblProjectManagerIssues = New System.Windows.Forms.Label
        Me.ColumnHeader25 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader24 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader23 = New System.Windows.Forms.ColumnHeader
        Me.ListView6 = New System.Windows.Forms.ListView
        Me.ColumnHeader34 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader35 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader36 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader37 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader38 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader47 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader26 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader27 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader39 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader40 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader41 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader42 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader43 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader44 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader45 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader46 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader69 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader70 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader71 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader72 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader73 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader74 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader75 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader76 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader77 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader78 = New System.Windows.Forms.ColumnHeader
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.lblOwnerLastEditedOn = New System.Windows.Forms.Label
        Me.lblOwnerLastEditedBy = New System.Windows.Forms.Label
        Me.btnRegister = New System.Windows.Forms.Button
        Me.pnlMain = New System.Windows.Forms.Panel
        Me.tbCntrlRegistration = New System.Windows.Forms.TabControl
        Me.tbPageOwnerDetail = New System.Windows.Forms.TabPage
        Me.pnlOwnerBottom = New System.Windows.Forms.Panel
        Me.tbCtrlOwner = New System.Windows.Forms.TabControl
        Me.tbPageOwnerFacilities = New System.Windows.Forms.TabPage
        Me.ugFacilityList = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerFacilityBottom = New System.Windows.Forms.Panel
        Me.lblNoOfFacilitiesValue = New System.Windows.Forms.Label
        Me.lblNoOfFacilities = New System.Windows.Forms.Label
        Me.tbPageOwnerContactList = New System.Windows.Forms.TabPage
        Me.pnlOwnerContactContainer = New System.Windows.Forms.Panel
        Me.ugOwnerContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlOwnerContactHeader = New System.Windows.Forms.Panel
        Me.chkOwnerShowActiveOnly = New System.Windows.Forms.CheckBox
        Me.chkOwnerShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkOwnerShowContactsforAllModules = New System.Windows.Forms.CheckBox
        Me.lblOwnerContacts = New System.Windows.Forms.Label
        Me.pnlOwnerContactButtons = New System.Windows.Forms.Panel
        Me.btnOwnerModifyContact = New System.Windows.Forms.Button
        Me.btnOwnerDeleteContact = New System.Windows.Forms.Button
        Me.btnOwnerAssociateContact = New System.Windows.Forms.Button
        Me.btnOwnerAddSearchContact = New System.Windows.Forms.Button
        Me.tbPrevFacs = New System.Windows.Forms.TabPage
        Me.ugPrevOwnedFacs = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlPrevOwnedFacsCount = New System.Windows.Forms.Panel
        Me.lblPreviousFacilitiesCountValue = New System.Windows.Forms.Label
        Me.lblNoofPreviousFacilities = New System.Windows.Forms.Label
        Me.tbPageOwnerDocuments = New System.Windows.Forms.TabPage
        Me.UCOwnerDocuments = New MUSTER.DocumentViewControl
        Me.pnlOwnerDetail = New System.Windows.Forms.Panel
        Me.btnLabels = New System.Windows.Forms.Button
        Me.btnEnvelopes = New System.Windows.Forms.Button
        Me.lblNewOwnerSnippetValue = New System.Windows.Forms.Label
        Me.chkOwnerAgencyInterest = New System.Windows.Forms.CheckBox
        Me.lblOwnerActiveOrNot = New System.Windows.Forms.Label
        Me.LinkLblCAPSignup = New System.Windows.Forms.LinkLabel
        Me.lblCAPParticipationLevel = New System.Windows.Forms.Label
        Me.pnlOwnerName = New System.Windows.Forms.Panel
        Me.pnlOwnerNameButton = New System.Windows.Forms.Panel
        Me.btnOwnerNameClose = New System.Windows.Forms.Button
        Me.btnOwnerNameCancel = New System.Windows.Forms.Button
        Me.btnOwnerNameOK = New System.Windows.Forms.Button
        Me.pnlOwnerOrg = New System.Windows.Forms.Panel
        Me.txtOwnerOrgName = New System.Windows.Forms.TextBox
        Me.lblOwnerOrgName = New System.Windows.Forms.Label
        Me.cmbOwnerOrgEntityCode = New System.Windows.Forms.ComboBox
        Me.lblOwnerOrgEntityCode = New System.Windows.Forms.Label
        Me.pnlPersonOrganization = New System.Windows.Forms.Panel
        Me.btnOwnerNameSearch = New System.Windows.Forms.Button
        Me.rdOwnerOrg = New System.Windows.Forms.RadioButton
        Me.rdOwnerPerson = New System.Windows.Forms.RadioButton
        Me.pnlOwnerPerson = New System.Windows.Forms.Panel
        Me.cmbOwnerNameSuffix = New System.Windows.Forms.ComboBox
        Me.cmbOwnerNameTitle = New System.Windows.Forms.ComboBox
        Me.lblOwnerNameSuffix = New System.Windows.Forms.Label
        Me.lblOwnerNameTitle = New System.Windows.Forms.Label
        Me.txtOwnerMiddleName = New System.Windows.Forms.TextBox
        Me.lblOwnerMiddleName = New System.Windows.Forms.Label
        Me.txtOwnerLastName = New System.Windows.Forms.TextBox
        Me.txtOwnerFirstName = New System.Windows.Forms.TextBox
        Me.lblOwnerLastName = New System.Windows.Forms.Label
        Me.lblOwnerFirstName = New System.Windows.Forms.Label
        Me.mskTxtOwnerFax = New AxMSMask.AxMaskEdBox
        Me.mskTxtOwnerPhone2 = New AxMSMask.AxMaskEdBox
        Me.mskTxtOwnerPhone = New AxMSMask.AxMaskEdBox
        Me.lblOwnerEmail = New System.Windows.Forms.Label
        Me.txtOwnerEmail = New System.Windows.Forms.TextBox
        Me.lblFax = New System.Windows.Forms.Label
        Me.pnlOwnerButtons = New System.Windows.Forms.Panel
        Me.btnOwnerFlag = New System.Windows.Forms.Button
        Me.btnOwnerComment = New System.Windows.Forms.Button
        Me.btnTransferOwnership = New System.Windows.Forms.Button
        Me.btnSaveOwner = New System.Windows.Forms.Button
        Me.btnDeleteOwner = New System.Windows.Forms.Button
        Me.btnOwnerCancel = New System.Windows.Forms.Button
        Me.txtOwnerAddress = New System.Windows.Forms.TextBox
        Me.lblOwnerAddress = New System.Windows.Forms.Label
        Me.txtOwnerName = New System.Windows.Forms.TextBox
        Me.lblOwnerName = New System.Windows.Forms.Label
        Me.lblOwnerStatus = New System.Windows.Forms.Label
        Me.lblOwnerCapParticipant = New System.Windows.Forms.Label
        Me.lblPhone2 = New System.Windows.Forms.Label
        Me.lblOwnerType = New System.Windows.Forms.Label
        Me.txtOwnerAIID = New System.Windows.Forms.TextBox
        Me.lblOwnerAIID = New System.Windows.Forms.Label
        Me.lblOwnerIDValue = New System.Windows.Forms.Label
        Me.lblOwnerID = New System.Windows.Forms.Label
        Me.lblOwnerPhone = New System.Windows.Forms.Label
        Me.cmbOwnerType = New System.Windows.Forms.ComboBox
        Me.chkCAPParticipant = New System.Windows.Forms.CheckBox
        Me.tbPageFacilityDetail = New System.Windows.Forms.TabPage
        Me.pnlFacilityBottom = New System.Windows.Forms.Panel
        Me.tabCtrlFacilityTankPipe = New System.Windows.Forms.TabControl
        Me.tbPageTankPipe = New System.Windows.Forms.TabPage
        Me.pnlTankPipe = New System.Windows.Forms.Panel
        Me.btnExpand = New System.Windows.Forms.Button
        Me.btnAddTank = New System.Windows.Forms.Button
        Me.dgPipesAndTanks = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label12 = New System.Windows.Forms.Label
        Me.pnlFacilityTankPipeButton = New System.Windows.Forms.Panel
        Me.lblTotalNoOfTanksValue = New System.Windows.Forms.Label
        Me.lblTotalNoOfTanks = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbPageFacilityContactList = New System.Windows.Forms.TabPage
        Me.pnlFacilityContactContainer = New System.Windows.Forms.Panel
        Me.ugFacilityContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label2 = New System.Windows.Forms.Label
        Me.pnlFacilityContactHeader = New System.Windows.Forms.Panel
        Me.chkFacilityShowActiveContactOnly = New System.Windows.Forms.CheckBox
        Me.chkFacilityShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkFacilityShowContactsforAllModule = New System.Windows.Forms.CheckBox
        Me.lblFacilityContacts = New System.Windows.Forms.Label
        Me.pnlFacilityContactBottom = New System.Windows.Forms.Panel
        Me.btnFacilityModifyContact = New System.Windows.Forms.Button
        Me.btnFacilityDeleteContact = New System.Windows.Forms.Button
        Me.btnFacilityAssociateContact = New System.Windows.Forms.Button
        Me.btnFacilityAddSearchContact = New System.Windows.Forms.Button
        Me.tpPreviouslyOwnedOwners = New System.Windows.Forms.TabPage
        Me.ugPreviouslyOwnedOwners = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.lblNoofOwnersValue = New System.Windows.Forms.Label
        Me.lblNoOfOwners = New System.Windows.Forms.Label
        Me.tbPageFacilityDocuments = New System.Windows.Forms.TabPage
        Me.UCFacilityDocuments = New MUSTER.DocumentViewControl
        Me.pnl_FacilityDetail = New System.Windows.Forms.Panel
        Me.dtPickAssess = New System.Windows.Forms.DateTimePicker
        Me.lblAssessDate = New System.Windows.Forms.Label
        Me.chkProhibition = New System.Windows.Forms.CheckBox
        Me.LinkLblCAPSignupFac = New System.Windows.Forms.LinkLabel
        Me.btnFacilityLabels = New System.Windows.Forms.Button
        Me.btnFacilityEnvelopes = New System.Windows.Forms.Button
        Me.lblFacilityCompany = New System.Windows.Forms.Label
        Me.txtFacilityCompany = New System.Windows.Forms.TextBox
        Me.lblFacilityLicensee = New System.Windows.Forms.Label
        Me.lblFacilityLicenseeSearch = New System.Windows.Forms.Label
        Me.txtFacilityLicensee = New System.Windows.Forms.TextBox
        Me.txtDueByNF = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.lblActiveLust = New System.Windows.Forms.Label
        Me.dtPickUpcomingInstallDateValue = New System.Windows.Forms.DateTimePicker
        Me.lblUpcomingInstallDate = New System.Windows.Forms.Label
        Me.chkUpcomingInstall = New System.Windows.Forms.CheckBox
        Me.lnkLblNextFac = New System.Windows.Forms.LinkLabel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnFacilityCancel = New System.Windows.Forms.Button
        Me.btnFacComments = New System.Windows.Forms.Button
        Me.btnDeleteFacility = New System.Windows.Forms.Button
        Me.btnFacilitySave = New System.Windows.Forms.Button
        Me.btnFacFlags = New System.Windows.Forms.Button
        Me.lblCAPStatusValue = New System.Windows.Forms.Label
        Me.lblCAPStatus = New System.Windows.Forms.Label
        Me.txtFuelBrand = New System.Windows.Forms.TextBox
        Me.ll = New System.Windows.Forms.Label
        Me.lblDateTransfered = New System.Windows.Forms.Label
        Me.dtFacilityPowerOff = New System.Windows.Forms.DateTimePicker
        Me.lnkLblPrevFacility = New System.Windows.Forms.LinkLabel
        Me.chkLUSTSite = New System.Windows.Forms.CheckBox
        Me.lblPowerOff = New System.Windows.Forms.Label
        Me.chkCAPCandidate = New System.Windows.Forms.CheckBox
        Me.lblFacilityLocationType = New System.Windows.Forms.Label
        Me.cmbFacilityLocationType = New System.Windows.Forms.ComboBox
        Me.lblFacilityMethod = New System.Windows.Forms.Label
        Me.cmbFacilityMethod = New System.Windows.Forms.ComboBox
        Me.lblFacilityDatum = New System.Windows.Forms.Label
        Me.cmbFacilityDatum = New System.Windows.Forms.ComboBox
        Me.cmbFacilityType = New System.Windows.Forms.ComboBox
        Me.txtFacilityLatSec = New System.Windows.Forms.TextBox
        Me.txtFacilityLongSec = New System.Windows.Forms.TextBox
        Me.txtFacilityLatMin = New System.Windows.Forms.TextBox
        Me.txtFacilityLongMin = New System.Windows.Forms.TextBox
        Me.lblFacilityLongMin = New System.Windows.Forms.Label
        Me.lblFacilityLongSec = New System.Windows.Forms.Label
        Me.lblFacilityLatMin = New System.Windows.Forms.Label
        Me.lblFacilityLatSec = New System.Windows.Forms.Label
        Me.lblFacilityLongDegree = New System.Windows.Forms.Label
        Me.mskTxtFacilityFax = New AxMSMask.AxMaskEdBox
        Me.mskTxtFacilityPhone = New AxMSMask.AxMaskEdBox
        Me.txtFacilityAddress = New System.Windows.Forms.TextBox
        Me.lblFacilitySIC = New System.Windows.Forms.Label
        Me.txtFacilityFax = New System.Windows.Forms.TextBox
        Me.lblFacilityFax = New System.Windows.Forms.Label
        Me.dtPickFacilityRecvd = New System.Windows.Forms.DateTimePicker
        Me.lblDateReceived = New System.Windows.Forms.Label
        Me.chkSignatureofNF = New System.Windows.Forms.CheckBox
        Me.lblFacilitySigOnNF = New System.Windows.Forms.Label
        Me.lblFacilityFuelBrand = New System.Windows.Forms.Label
        Me.lblFacilityStatusValue = New System.Windows.Forms.Label
        Me.lblFacilityStatus = New System.Windows.Forms.Label
        Me.txtFacilityLongDegree = New System.Windows.Forms.TextBox
        Me.txtFacilityLatDegree = New System.Windows.Forms.TextBox
        Me.lblFacilityLongitude = New System.Windows.Forms.Label
        Me.lblFacilityLatitude = New System.Windows.Forms.Label
        Me.lblFacilityType = New System.Windows.Forms.Label
        Me.txtFacilityAIID = New System.Windows.Forms.TextBox
        Me.lblfacilityAIID = New System.Windows.Forms.Label
        Me.lblFacilityIDValue = New System.Windows.Forms.Label
        Me.lblFacilityID = New System.Windows.Forms.Label
        Me.txtfacilityPhone = New System.Windows.Forms.TextBox
        Me.lblFacilityPhone = New System.Windows.Forms.Label
        Me.txtFacilityName = New System.Windows.Forms.TextBox
        Me.lblFacilityAddress = New System.Windows.Forms.Label
        Me.lblFacilityName = New System.Windows.Forms.Label
        Me.txtFacilityZip = New System.Windows.Forms.TextBox
        Me.lblFacilityLatDegree = New System.Windows.Forms.Label
        Me.txtFacilitySIC = New System.Windows.Forms.Label
        Me.txtFacilityNameForEnsite = New System.Windows.Forms.TextBox
        Me.lblFacilityNameForEnsite = New System.Windows.Forms.Label
        Me.tbPageManageTank = New System.Windows.Forms.TabPage
        Me.dgPipesAndTanks2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.tbCntrlPipe = New System.Windows.Forms.TabControl
        Me.tbPagePipeDetail = New System.Windows.Forms.TabPage
        Me.pnlPipeDetail = New System.Windows.Forms.Panel
        Me.pnlPipeClosure = New System.Windows.Forms.Panel
        Me.cmbPipeInertFill = New System.Windows.Forms.ComboBox
        Me.cmbPipeClosureType = New System.Windows.Forms.ComboBox
        Me.dtPickPipeLastUsed = New System.Windows.Forms.DateTimePicker
        Me.lblPipeLastUsed = New System.Windows.Forms.Label
        Me.lblPipeClosedOn = New System.Windows.Forms.Label
        Me.lblPipeClosedOnDate = New System.Windows.Forms.Label
        Me.lblPipeInertFillValue = New System.Windows.Forms.Label
        Me.lblPipeInertFill = New System.Windows.Forms.Label
        Me.lblPipeClosureStatusValue = New System.Windows.Forms.Label
        Me.lblPipeClosureStatus = New System.Windows.Forms.Label
        Me.lblPipeClosureRcvdDateValue = New System.Windows.Forms.Label
        Me.lblPipeClosureRcvd = New System.Windows.Forms.Label
        Me.Panel20 = New System.Windows.Forms.Panel
        Me.lblPipeClosureCaption = New System.Windows.Forms.Label
        Me.lblPipeClosure = New System.Windows.Forms.Label
        Me.pnlPipeInstallerOath = New System.Windows.Forms.Panel
        Me.txtPipeLicensee = New System.Windows.Forms.TextBox
        Me.lblPipeLicenseeCompanySearch = New System.Windows.Forms.Label
        Me.txtPipeCompanyName = New System.Windows.Forms.TextBox
        Me.lblPipeInstallerName = New System.Windows.Forms.Label
        Me.lblPipeInstallerCompanyName = New System.Windows.Forms.Label
        Me.dtPickPipeSigned = New System.Windows.Forms.DateTimePicker
        Me.lblDtPipeSigned = New System.Windows.Forms.Label
        Me.Panel9 = New System.Windows.Forms.Panel
        Me.lblPipeInstallerOath = New System.Windows.Forms.Label
        Me.lblPipeInstallerOathDisplay = New System.Windows.Forms.Label
        Me.pnlPipeRelease = New System.Windows.Forms.Panel
        Me.dtElectronicDeviceInspected = New System.Windows.Forms.DateTimePicker
        Me.LblElectronicDeviceInspected = New System.Windows.Forms.Label
        Me.grpBxReleaseDetectionGroup2 = New System.Windows.Forms.GroupBox
        Me.dtPickPipeLeakDetectorTest = New System.Windows.Forms.DateTimePicker
        Me.lblPipeReleaseDetection2 = New System.Windows.Forms.Label
        Me.cmbPipeReleaseDetection2 = New System.Windows.Forms.ComboBox
        Me.lblPipeALLDTestDate = New System.Windows.Forms.Label
        Me.grpBxPipeReleaseDetectionGroup1 = New System.Windows.Forms.GroupBox
        Me.dtPickPipeTightnessTest = New System.Windows.Forms.DateTimePicker
        Me.lblPipeReleaseDetection = New System.Windows.Forms.Label
        Me.cmbPipeReleaseDetection1 = New System.Windows.Forms.ComboBox
        Me.lblLPipeTightnessTstDt = New System.Windows.Forms.Label
        Me.LblSecondaryContainmentInspected = New System.Windows.Forms.Label
        Me.dtSecondaryContainmentInspected = New System.Windows.Forms.DateTimePicker
        Me.Panel11 = New System.Windows.Forms.Panel
        Me.lblPipeReleaseHead = New System.Windows.Forms.Label
        Me.lblPipeReleaseDisplay = New System.Windows.Forms.Label
        Me.pnlPipeType = New System.Windows.Forms.Panel
        Me.dtSheerValueTest = New System.Windows.Forms.DateTimePicker
        Me.LblSheerValueTestDate = New System.Windows.Forms.Label
        Me.cmbPipeType = New System.Windows.Forms.ComboBox
        Me.lblPipeCapacity = New System.Windows.Forms.Label
        Me.Panel15 = New System.Windows.Forms.Panel
        Me.lblPipeTypeCaption = New System.Windows.Forms.Label
        Me.lblPipeType = New System.Windows.Forms.Label
        Me.pnlPipeMaterial = New System.Windows.Forms.Panel
        Me.lblEmergencyPowerPipe = New System.Windows.Forms.Label
        Me.chkEmergencyPowerPipe = New System.Windows.Forms.CheckBox
        Me.lblPipeManufacturer = New System.Windows.Forms.Label
        Me.cmbPipeManufacturerID = New System.Windows.Forms.ComboBox
        Me.dtPickPipeTerminationCPLastTested = New System.Windows.Forms.DateTimePicker
        Me.lblPipeTerminationCPLastTest = New System.Windows.Forms.Label
        Me.dtPickPipeTerminationCPInstalled = New System.Windows.Forms.DateTimePicker
        Me.lblPipeTerminationCPInstalled = New System.Windows.Forms.Label
        Me.grpBxPipeTerminationAtDispenser = New System.Windows.Forms.GroupBox
        Me.cmbPipeTerminationDispenserCPType = New System.Windows.Forms.ComboBox
        Me.cmbPipeTerminationDispenserType = New System.Windows.Forms.ComboBox
        Me.lblPipeTerminationDispenserCPType = New System.Windows.Forms.Label
        Me.lblPipeTerminationDispenserType = New System.Windows.Forms.Label
        Me.grpBxPipeTerminationAtTank = New System.Windows.Forms.GroupBox
        Me.cmbPipeTerminationTankCPType = New System.Windows.Forms.ComboBox
        Me.cmbPipeTerminationTankType = New System.Windows.Forms.ComboBox
        Me.lblPipeTerminationTankCPType = New System.Windows.Forms.Label
        Me.lblPipeTerminationTankType = New System.Windows.Forms.Label
        Me.grpBxPipeContainmentSumpsLocation = New System.Windows.Forms.GroupBox
        Me.chkPipeSumpsAtTank = New System.Windows.Forms.CheckBox
        Me.chkPipeSumpsAtDispenser = New System.Windows.Forms.CheckBox
        Me.dtPickPipeCPLastTest = New System.Windows.Forms.DateTimePicker
        Me.dtPickPipeCPInstalled = New System.Windows.Forms.DateTimePicker
        Me.lblPipeMaterial = New System.Windows.Forms.Label
        Me.lblPipeOption = New System.Windows.Forms.Label
        Me.cmbPipeMaterial = New System.Windows.Forms.ComboBox
        Me.cmbPipeOptions = New System.Windows.Forms.ComboBox
        Me.lblPipeCPType = New System.Windows.Forms.Label
        Me.cmbPipeCPType = New System.Windows.Forms.ComboBox
        Me.lblDatePipeCPInstalled = New System.Windows.Forms.Label
        Me.lblDtPipeLastTested = New System.Windows.Forms.Label
        Me.pnlPipeMaterialHead = New System.Windows.Forms.Panel
        Me.lblPipeMaterialHead = New System.Windows.Forms.Label
        Me.lblPipeMaterialDisplay = New System.Windows.Forms.Label
        Me.pnlPipeDateOfInstallation = New System.Windows.Forms.Panel
        Me.lblPipeFuelTypeValue = New System.Windows.Forms.Label
        Me.lblPipeCerclaNoValue = New System.Windows.Forms.Label
        Me.lblPipeCerclaNo = New System.Windows.Forms.Label
        Me.lblPipeSubstance = New System.Windows.Forms.Label
        Me.lblPipeSubstanceValue = New System.Windows.Forms.Label
        Me.dtPickDatePipePlacedInService = New System.Windows.Forms.DateTimePicker
        Me.lblPipePlacedInServiceOn = New System.Windows.Forms.Label
        Me.lblPipeInstallationPlanedFor = New System.Windows.Forms.Label
        Me.dtPickPipePlannedInstallation = New System.Windows.Forms.DateTimePicker
        Me.lblPipeInstalledOn = New System.Windows.Forms.Label
        Me.dtPickPipeInstalled = New System.Windows.Forms.DateTimePicker
        Me.lblPipeFuelType = New System.Windows.Forms.Label
        Me.Panel13 = New System.Windows.Forms.Panel
        Me.lblPipeDateOfInstallationCaption = New System.Windows.Forms.Label
        Me.lblPipeDateOfInstallation = New System.Windows.Forms.Label
        Me.pnlPipeDescription = New System.Windows.Forms.Panel
        Me.cmbPipeStatus = New System.Windows.Forms.ComboBox
        Me.lblPipeStatus = New System.Windows.Forms.Label
        Me.pnlPipeDescHead = New System.Windows.Forms.Panel
        Me.lblPipeDescHead = New System.Windows.Forms.Label
        Me.lblPipeDescDisplay = New System.Windows.Forms.Label
        Me.pnlPipeButtons = New System.Windows.Forms.Panel
        Me.btnPipeCancel = New System.Windows.Forms.Button
        Me.btnPipeComments = New System.Windows.Forms.Button
        Me.btnToTank = New System.Windows.Forms.Button
        Me.btnCopyPipeProfile = New System.Windows.Forms.Button
        Me.btnDeletePipe = New System.Windows.Forms.Button
        Me.btnPipeSave = New System.Windows.Forms.Button
        Me.pnlPipeDetailHeader = New System.Windows.Forms.Panel
        Me.lblPipeTankIDValue = New System.Windows.Forms.Label
        Me.lblPipeIDValue = New System.Windows.Forms.Label
        Me.lblPipeIndex = New System.Windows.Forms.Label
        Me.lblPipeCompartmentIndex = New System.Windows.Forms.Label
        Me.lblPipeCompartment = New System.Windows.Forms.Label
        Me.lblPipeID = New System.Windows.Forms.Label
        Me.lblPipeTankID = New System.Windows.Forms.Label
        Me.lblTankCountVal2 = New System.Windows.Forms.Label
        Me.lblTankIDValue2 = New System.Windows.Forms.Label
        Me.tbCntrlTank = New System.Windows.Forms.TabControl
        Me.tbPageTankDetail = New System.Windows.Forms.TabPage
        Me.pnlTankDetail = New System.Windows.Forms.Panel
        Me.pnlTankClosure = New System.Windows.Forms.Panel
        Me.cmbTankInertFill = New System.Windows.Forms.ComboBox
        Me.cmbTankClosureType = New System.Windows.Forms.ComboBox
        Me.lblTankInertFillValue = New System.Windows.Forms.Label
        Me.lblTankInertFill = New System.Windows.Forms.Label
        Me.lblTankClosureStatusValue = New System.Windows.Forms.Label
        Me.lblTankClosureStatus = New System.Windows.Forms.Label
        Me.lblDateTankClosureRecvValue = New System.Windows.Forms.Label
        Me.lblDateClosureRecvd = New System.Windows.Forms.Label
        Me.lblDateClosed = New System.Windows.Forms.Label
        Me.dtPickLastUsed = New System.Windows.Forms.DateTimePicker
        Me.lblDtLastUsed = New System.Windows.Forms.Label
        Me.lblClosuredate = New System.Windows.Forms.Label
        Me.pnllblTankClosure = New System.Windows.Forms.Panel
        Me.lblTankClosureCaption = New System.Windows.Forms.Label
        Me.lblTankClosure = New System.Windows.Forms.Label
        Me.pnlTankInstallerOath = New System.Windows.Forms.Panel
        Me.txtLicensee = New System.Windows.Forms.TextBox
        Me.lblLicenseeSearch = New System.Windows.Forms.Label
        Me.txtTankCompany = New System.Windows.Forms.TextBox
        Me.lblLicenseeName = New System.Windows.Forms.Label
        Me.lblTankInstallerCompany = New System.Windows.Forms.Label
        Me.dtPickTankInstallerSigned = New System.Windows.Forms.DateTimePicker
        Me.lblTankInstallerDtSigned = New System.Windows.Forms.Label
        Me.pnlInstallerOath = New System.Windows.Forms.Panel
        Me.lblTankInstallerOath = New System.Windows.Forms.Label
        Me.lblTankInstallerOathDisplay = New System.Windows.Forms.Label
        Me.pnlTankRelease = New System.Windows.Forms.Panel
        Me.dtPickATGLastInspected = New System.Windows.Forms.DateTimePicker
        Me.dtPickElectronicDeviceInspected = New System.Windows.Forms.DateTimePicker
        Me.LblDateATGLastInspected = New System.Windows.Forms.Label
        Me.LblDateElectronicDeviceInspected = New System.Windows.Forms.Label
        Me.cmbTankReleaseDetection = New System.Windows.Forms.ComboBox
        Me.dtPickTankTightnessTest = New System.Windows.Forms.DateTimePicker
        Me.lblRelseDetection = New System.Windows.Forms.Label
        Me.lblLTankTightnessTstDt = New System.Windows.Forms.Label
        Me.chkTankDrpTubeInvControl = New System.Windows.Forms.CheckBox
        Me.dtPickSecondaryContainmentLastInspected = New System.Windows.Forms.DateTimePicker
        Me.lbLDateSecondaryContainmentLastInspected = New System.Windows.Forms.Label
        Me.pnlReleaseDetection = New System.Windows.Forms.Panel
        Me.lblTankReleaseHead = New System.Windows.Forms.Label
        Me.lblTankReleaseDisplay = New System.Windows.Forms.Label
        Me.pnlTankMaterial = New System.Windows.Forms.Panel
        Me.dtPickOverfillPreventionInstalled = New System.Windows.Forms.DateTimePicker
        Me.LblDateOverfillPreventionInstalled = New System.Windows.Forms.Label
        Me.dtPickOverfillPreventionLastInspected = New System.Windows.Forms.DateTimePicker
        Me.dtPickSpillPreventionLastTested = New System.Windows.Forms.DateTimePicker
        Me.LblDateOverfillPreventionLastInspected = New System.Windows.Forms.Label
        Me.LblDateSpillPreventionLastTested = New System.Windows.Forms.Label
        Me.dtPickSpillPreventionInstalled = New System.Windows.Forms.DateTimePicker
        Me.LblDateSpillPreventionInstalled = New System.Windows.Forms.Label
        Me.chkDeliveriesLimited = New System.Windows.Forms.CheckBox
        Me.dtPickCPLastTested = New System.Windows.Forms.DateTimePicker
        Me.dtPickCPInstalled = New System.Windows.Forms.DateTimePicker
        Me.cmbTankOverfillProtectionType = New System.Windows.Forms.ComboBox
        Me.dtPickInteriorLiningInstalled = New System.Windows.Forms.DateTimePicker
        Me.dtPickLastInteriorLinningInspection = New System.Windows.Forms.DateTimePicker
        Me.lblTankMaterial = New System.Windows.Forms.Label
        Me.lblTankOption = New System.Windows.Forms.Label
        Me.cmbTankMaterial = New System.Windows.Forms.ComboBox
        Me.cmbTankOptions = New System.Windows.Forms.ComboBox
        Me.lblTankCPType = New System.Windows.Forms.Label
        Me.cmbTankCPType = New System.Windows.Forms.ComboBox
        Me.lblDateCPInstalled = New System.Windows.Forms.Label
        Me.lblDtTankLstTested = New System.Windows.Forms.Label
        Me.chkOverFilledProtected = New System.Windows.Forms.CheckBox
        Me.chkBxSpillProtected = New System.Windows.Forms.CheckBox
        Me.chkBxTightfillAdapters = New System.Windows.Forms.CheckBox
        Me.lblDateLnInteriorInstalled = New System.Windows.Forms.Label
        Me.lblDtLnInteriorLstInspect = New System.Windows.Forms.Label
        Me.lblTankOverfillProtectionType = New System.Windows.Forms.Label
        Me.chkEmergencyPower = New System.Windows.Forms.CheckBox
        Me.pnlTankMaterialHead = New System.Windows.Forms.Panel
        Me.lblTankMaterialHead = New System.Windows.Forms.Label
        Me.lblTankMaterialDisplay = New System.Windows.Forms.Label
        Me.pnlTankTotalCapacity = New System.Windows.Forms.Panel
        Me.txtTankCapacity = New System.Windows.Forms.Label
        Me.txtTankCompartmentNumber = New System.Windows.Forms.Label
        Me.pnlNonCompProperties = New System.Windows.Forms.Panel
        Me.lblCERCLAtt = New System.Windows.Forms.Label
        Me.lblTankManifoldValue = New System.Windows.Forms.Label
        Me.cmbTankCercla = New System.Windows.Forms.ComboBox
        Me.cmbTankFuelType = New System.Windows.Forms.ComboBox
        Me.cmbTanksubstance = New System.Windows.Forms.ComboBox
        Me.lblNonCompTankCapacity = New System.Windows.Forms.Label
        Me.lblTankCercla = New System.Windows.Forms.Label
        Me.lblTankFuelType = New System.Windows.Forms.Label
        Me.txtNonCompTankCapacity = New System.Windows.Forms.TextBox
        Me.lblTankSubstance = New System.Windows.Forms.Label
        Me.lblTankManifold = New System.Windows.Forms.Label
        Me.cmbTankCerclaDesc = New System.Windows.Forms.ComboBox
        Me.dGridCompartments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lblTankType = New System.Windows.Forms.Label
        Me.cmbTankType = New System.Windows.Forms.ComboBox
        Me.cmbTankCompCercla = New System.Windows.Forms.ComboBox
        Me.cmbTankCompSubstance = New System.Windows.Forms.ComboBox
        Me.lblTankCapacity = New System.Windows.Forms.Label
        Me.chkTankCompartment = New System.Windows.Forms.CheckBox
        Me.lblTankCompartmentNumber = New System.Windows.Forms.Label
        Me.Panel8 = New System.Windows.Forms.Panel
        Me.pnllblTankTotalCapacity = New System.Windows.Forms.Panel
        Me.lblTankTotalCapcityCaption = New System.Windows.Forms.Label
        Me.lblTankTotalCapacity = New System.Windows.Forms.Label
        Me.pnlTankInstallation = New System.Windows.Forms.Panel
        Me.lblTankInstalledOn = New System.Windows.Forms.Label
        Me.cmbTankManufacturer = New System.Windows.Forms.ComboBox
        Me.lblTankManufacturer = New System.Windows.Forms.Label
        Me.dtPickDatePlacedInService = New System.Windows.Forms.DateTimePicker
        Me.lblTankPlacedInServiceOn = New System.Windows.Forms.Label
        Me.dtPickTankInstalled = New System.Windows.Forms.DateTimePicker
        Me.dtPickPlannedInstallation = New System.Windows.Forms.DateTimePicker
        Me.lblTankInstallationPlanedFor = New System.Windows.Forms.Label
        Me.pnllblDateofInstallation = New System.Windows.Forms.Panel
        Me.lblDateofInstallation = New System.Windows.Forms.Label
        Me.lblTankInstallation = New System.Windows.Forms.Label
        Me.pnlTankDescriptionTop = New System.Windows.Forms.Panel
        Me.chkTankProhibition = New System.Windows.Forms.CheckBox
        Me.chkBoxReplacementTank = New System.Windows.Forms.CheckBox
        Me.lblTankOriginalStatusID = New System.Windows.Forms.Label
        Me.cmbTankStatus = New System.Windows.Forms.ComboBox
        Me.lblTankStatus = New System.Windows.Forms.Label
        Me.txtTankFacilityID = New System.Windows.Forms.TextBox
        Me.pnlTankDescHead = New System.Windows.Forms.Panel
        Me.lblTankDescHead = New System.Windows.Forms.Label
        Me.lblTankDescDisplay = New System.Windows.Forms.Label
        Me.pnlTankDetailHeader = New System.Windows.Forms.Panel
        Me.lblTankCountVal = New System.Windows.Forms.Label
        Me.btnAddTank2 = New System.Windows.Forms.Button
        Me.lblTankIDValue = New System.Windows.Forms.Label
        Me.lblTankID = New System.Windows.Forms.Label
        Me.btnAddPipe = New System.Windows.Forms.Button
        Me.btnAddExistingPipe = New System.Windows.Forms.Button
        Me.btnDetachPipes = New System.Windows.Forms.Button
        Me.pnlTankButtons = New System.Windows.Forms.Panel
        Me.btnTankCancel = New System.Windows.Forms.Button
        Me.btnTankComments = New System.Windows.Forms.Button
        Me.btnToPipe = New System.Windows.Forms.Button
        Me.btnCopyTankProfileToNew = New System.Windows.Forms.Button
        Me.btnDeleteTank = New System.Windows.Forms.Button
        Me.btnTankSave = New System.Windows.Forms.Button
        Me.pnlTankDetailMainDisplay = New System.Windows.Forms.Panel
        Me.btnExpandTP2 = New System.Windows.Forms.Button
        Me.lnkLblNextTank = New System.Windows.Forms.LinkLabel
        Me.lnkLblPrevTank = New System.Windows.Forms.LinkLabel
        Me.pnlTankCount2 = New System.Windows.Forms.Panel
        Me.lblTotalNoOfTanksValue2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbPageSummary = New System.Windows.Forms.TabPage
        Me.pnlOwnerSummaryDetails = New System.Windows.Forms.Panel
        Me.UCOwnerSummary = New MUSTER.OwnerSummary
        Me.Panel12 = New System.Windows.Forms.Panel
        Me.pnlOwnerSummaryHeader = New System.Windows.Forms.Panel
        Me.ctxMenuTankCompartment = New System.Windows.Forms.ContextMenu
        Me.mnuAddCompPipe = New System.Windows.Forms.MenuItem
        Me.mnuShowCompPipes = New System.Windows.Forms.MenuItem
        Me.ctxMenuTankPipe = New System.Windows.Forms.ContextMenu
        Me.MI_DeleteTankPipe = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MI_NewTank1 = New System.Windows.Forms.MenuItem
        Me.MI_EditTank1 = New System.Windows.Forms.MenuItem
        Me.LI_AttachPipes1 = New System.Windows.Forms.MenuItem
        Me.MI_CopyTank1 = New System.Windows.Forms.MenuItem
        Me.LI_AddTankCompartment1 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MI_EditPipe = New System.Windows.Forms.MenuItem
        Me.MI_CopyPipe = New System.Windows.Forms.MenuItem
        Me.MI_DetachPipes = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ctxMenuTank = New System.Windows.Forms.ContextMenu
        Me.MI_NewTank = New System.Windows.Forms.MenuItem
        Me.MI_EditTank = New System.Windows.Forms.MenuItem
        Me.MI_CopyTank = New System.Windows.Forms.MenuItem
        Me.LI_AddTankCompartment = New System.Windows.Forms.MenuItem
        Me.MI_DeleteTank = New System.Windows.Forms.MenuItem
        Me.MI_AttachPipes = New System.Windows.Forms.MenuItem
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.DateTimePicker8 = New System.Windows.Forms.DateTimePicker
        Me.DateTimePicker9 = New System.Windows.Forms.DateTimePicker
        Me.CheckBox10 = New System.Windows.Forms.CheckBox
        Me.Panel16 = New System.Windows.Forms.Panel
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.DateTimePicker10 = New System.Windows.Forms.DateTimePicker
        Me.DateTimePicker11 = New System.Windows.Forms.DateTimePicker
        Me.CheckBox11 = New System.Windows.Forms.CheckBox
        Me.Panel17 = New System.Windows.Forms.Panel
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.DateTimePicker12 = New System.Windows.Forms.DateTimePicker
        Me.DateTimePicker13 = New System.Windows.Forms.DateTimePicker
        Me.CheckBox12 = New System.Windows.Forms.CheckBox
        Me.Panel18 = New System.Windows.Forms.Panel
        Me.DateTimePicker14 = New System.Windows.Forms.DateTimePicker
        Me.DateTimePicker15 = New System.Windows.Forms.DateTimePicker
        Me.CheckBox13 = New System.Windows.Forms.CheckBox
        Me.Panel19 = New System.Windows.Forms.Panel
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.Button9 = New System.Windows.Forms.Button
        Me.btnTankCompCol8 = New System.Windows.Forms.Button
        Me.btnTankCompCol7 = New System.Windows.Forms.Button
        Me.btnTankCompCol6 = New System.Windows.Forms.Button
        Me.btnTankCompCol5 = New System.Windows.Forms.Button
        Me.btnTankCompCol4 = New System.Windows.Forms.Button
        Me.btnTankCompCol3 = New System.Windows.Forms.Button
        Me.btnTankCompCol2 = New System.Windows.Forms.Button
        Me.btnTankCompCol1 = New System.Windows.Forms.Button
        Me.pnlTankCompartmentHeader = New System.Windows.Forms.Panel
        Me.lblNoOfPreviouslyOwnedFacilitiesValue = New System.Windows.Forms.Label
        Me.lblNoOfPreviouslyOwnedFacilities = New System.Windows.Forms.Label
        Me.pnlTop.SuspendLayout()
        Me.pnlMain.SuspendLayout()
        Me.tbCntrlRegistration.SuspendLayout()
        Me.tbPageOwnerDetail.SuspendLayout()
        Me.pnlOwnerBottom.SuspendLayout()
        Me.tbCtrlOwner.SuspendLayout()
        Me.tbPageOwnerFacilities.SuspendLayout()
        CType(Me.ugFacilityList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerFacilityBottom.SuspendLayout()
        Me.tbPageOwnerContactList.SuspendLayout()
        Me.pnlOwnerContactContainer.SuspendLayout()
        CType(Me.ugOwnerContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerContactHeader.SuspendLayout()
        Me.pnlOwnerContactButtons.SuspendLayout()
        Me.tbPrevFacs.SuspendLayout()
        CType(Me.ugPrevOwnedFacs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPrevOwnedFacsCount.SuspendLayout()
        Me.tbPageOwnerDocuments.SuspendLayout()
        Me.pnlOwnerDetail.SuspendLayout()
        Me.pnlOwnerName.SuspendLayout()
        Me.pnlOwnerNameButton.SuspendLayout()
        Me.pnlOwnerOrg.SuspendLayout()
        Me.pnlPersonOrganization.SuspendLayout()
        Me.pnlOwnerPerson.SuspendLayout()
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerButtons.SuspendLayout()
        Me.tbPageFacilityDetail.SuspendLayout()
        Me.pnlFacilityBottom.SuspendLayout()
        Me.tabCtrlFacilityTankPipe.SuspendLayout()
        Me.tbPageTankPipe.SuspendLayout()
        Me.pnlTankPipe.SuspendLayout()
        CType(Me.dgPipesAndTanks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityTankPipeButton.SuspendLayout()
        Me.tbPageFacilityContactList.SuspendLayout()
        Me.pnlFacilityContactContainer.SuspendLayout()
        CType(Me.ugFacilityContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityContactHeader.SuspendLayout()
        Me.pnlFacilityContactBottom.SuspendLayout()
        Me.tpPreviouslyOwnedOwners.SuspendLayout()
        CType(Me.ugPreviouslyOwnedOwners, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.tbPageFacilityDocuments.SuspendLayout()
        Me.pnl_FacilityDetail.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageManageTank.SuspendLayout()
        CType(Me.dgPipesAndTanks2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbCntrlPipe.SuspendLayout()
        Me.tbPagePipeDetail.SuspendLayout()
        Me.pnlPipeDetail.SuspendLayout()
        Me.pnlPipeClosure.SuspendLayout()
        Me.Panel20.SuspendLayout()
        Me.pnlPipeInstallerOath.SuspendLayout()
        Me.Panel9.SuspendLayout()
        Me.pnlPipeRelease.SuspendLayout()
        Me.grpBxReleaseDetectionGroup2.SuspendLayout()
        Me.grpBxPipeReleaseDetectionGroup1.SuspendLayout()
        Me.Panel11.SuspendLayout()
        Me.pnlPipeType.SuspendLayout()
        Me.Panel15.SuspendLayout()
        Me.pnlPipeMaterial.SuspendLayout()
        Me.grpBxPipeTerminationAtDispenser.SuspendLayout()
        Me.grpBxPipeTerminationAtTank.SuspendLayout()
        Me.grpBxPipeContainmentSumpsLocation.SuspendLayout()
        Me.pnlPipeMaterialHead.SuspendLayout()
        Me.pnlPipeDateOfInstallation.SuspendLayout()
        Me.Panel13.SuspendLayout()
        Me.pnlPipeDescription.SuspendLayout()
        Me.pnlPipeDescHead.SuspendLayout()
        Me.pnlPipeButtons.SuspendLayout()
        Me.pnlPipeDetailHeader.SuspendLayout()
        Me.tbCntrlTank.SuspendLayout()
        Me.tbPageTankDetail.SuspendLayout()
        Me.pnlTankDetail.SuspendLayout()
        Me.pnlTankClosure.SuspendLayout()
        Me.pnllblTankClosure.SuspendLayout()
        Me.pnlTankInstallerOath.SuspendLayout()
        Me.pnlInstallerOath.SuspendLayout()
        Me.pnlTankRelease.SuspendLayout()
        Me.pnlReleaseDetection.SuspendLayout()
        Me.pnlTankMaterial.SuspendLayout()
        Me.pnlTankMaterialHead.SuspendLayout()
        Me.pnlTankTotalCapacity.SuspendLayout()
        Me.pnlNonCompProperties.SuspendLayout()
        CType(Me.dGridCompartments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel8.SuspendLayout()
        Me.pnllblTankTotalCapacity.SuspendLayout()
        Me.pnlTankInstallation.SuspendLayout()
        Me.pnllblDateofInstallation.SuspendLayout()
        Me.pnlTankDescriptionTop.SuspendLayout()
        Me.pnlTankDescHead.SuspendLayout()
        Me.pnlTankDetailHeader.SuspendLayout()
        Me.pnlTankButtons.SuspendLayout()
        Me.pnlTankDetailMainDisplay.SuspendLayout()
        Me.pnlTankCount2.SuspendLayout()
        Me.tbPageSummary.SuspendLayout()
        Me.pnlOwnerSummaryDetails.SuspendLayout()
        Me.Panel19.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbPageOwnerDocuments
        '
        Me.tbPageOwnerDocuments.Controls.Add(Me.UCOwnerDocuments)
        Me.tbPageOwnerDocuments.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerDocuments.Name = "tbPageOwnerDocuments"
        Me.tbPageOwnerDocuments.Size = New System.Drawing.Size(954, 331)
        Me.tbPageOwnerDocuments.TabIndex = 3
        Me.tbPageOwnerDocuments.Text = "Documents"
        '
        'UCOwnerDocuments
        '
        Me.UCOwnerDocuments.AutoScroll = True
        Me.UCOwnerDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerDocuments.Name = "UCOwnerDocuments"
        Me.UCOwnerDocuments.Size = New System.Drawing.Size(954, 331)
        Me.UCOwnerDocuments.TabIndex = 0
        '
        'pnlOwnerDetail
        '
        Me.pnlOwnerDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnlOwnerDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlOwnerDetail.Controls.Add(Me.btnLabels)
        Me.pnlOwnerDetail.Controls.Add(Me.btnEnvelopes)
        Me.pnlOwnerDetail.Controls.Add(Me.lblNewOwnerSnippetValue)
        Me.pnlOwnerDetail.Controls.Add(Me.chkOwnerAgencyInterest)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerActiveOrNot)
        Me.pnlOwnerDetail.Controls.Add(Me.LinkLblCAPSignup)
        Me.pnlOwnerDetail.Controls.Add(Me.lblCAPParticipationLevel)
        Me.pnlOwnerDetail.Controls.Add(Me.pnlOwnerName)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerFax)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerPhone2)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerPhone)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerEmail)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerEmail)
        Me.pnlOwnerDetail.Controls.Add(Me.lblFax)
        Me.pnlOwnerDetail.Controls.Add(Me.pnlOwnerButtons)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerAddress)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerAddress)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerName)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerName)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerStatus)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerCapParticipant)
        Me.pnlOwnerDetail.Controls.Add(Me.lblPhone2)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerType)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerAIID)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerAIID)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerIDValue)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerID)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerPhone)
        Me.pnlOwnerDetail.Controls.Add(Me.cmbOwnerType)
        Me.pnlOwnerDetail.Controls.Add(Me.chkCAPParticipant)
        Me.pnlOwnerDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerDetail.Name = "pnlOwnerDetail"
        Me.pnlOwnerDetail.Size = New System.Drawing.Size(964, 272)
        Me.pnlOwnerDetail.TabIndex = 0
        '
        'btnLabels
        '
        Me.btnLabels.Location = New System.Drawing.Point(2, 112)
        Me.btnLabels.Name = "btnLabels"
        Me.btnLabels.TabIndex = 1011
        Me.btnLabels.Text = "Labels"
        '
        'btnEnvelopes
        '
        Me.btnEnvelopes.Location = New System.Drawing.Point(2, 80)
        Me.btnEnvelopes.Name = "btnEnvelopes"
        Me.btnEnvelopes.TabIndex = 1010
        Me.btnEnvelopes.Text = "Envelopes"
        '
        'lblNewOwnerSnippetValue
        '
        Me.lblNewOwnerSnippetValue.Location = New System.Drawing.Point(848, 8)
        Me.lblNewOwnerSnippetValue.Name = "lblNewOwnerSnippetValue"
        Me.lblNewOwnerSnippetValue.Size = New System.Drawing.Size(8, 16)
        Me.lblNewOwnerSnippetValue.TabIndex = 1009
        Me.lblNewOwnerSnippetValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblNewOwnerSnippetValue.Visible = False
        '
        'chkOwnerAgencyInterest
        '
        Me.chkOwnerAgencyInterest.Location = New System.Drawing.Point(560, 16)
        Me.chkOwnerAgencyInterest.Name = "chkOwnerAgencyInterest"
        Me.chkOwnerAgencyInterest.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkOwnerAgencyInterest.Size = New System.Drawing.Size(120, 20)
        Me.chkOwnerAgencyInterest.TabIndex = 8
        Me.chkOwnerAgencyInterest.Text = "Agency Interest   "
        Me.chkOwnerAgencyInterest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOwnerActiveOrNot
        '
        Me.lblOwnerActiveOrNot.BackColor = System.Drawing.SystemColors.Control
        Me.lblOwnerActiveOrNot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOwnerActiveOrNot.Location = New System.Drawing.Point(424, 16)
        Me.lblOwnerActiveOrNot.Name = "lblOwnerActiveOrNot"
        Me.lblOwnerActiveOrNot.Size = New System.Drawing.Size(96, 20)
        Me.lblOwnerActiveOrNot.TabIndex = 1006
        '
        'LinkLblCAPSignup
        '
        Me.LinkLblCAPSignup.Location = New System.Drawing.Point(552, 74)
        Me.LinkLblCAPSignup.Name = "LinkLblCAPSignup"
        Me.LinkLblCAPSignup.Size = New System.Drawing.Size(152, 16)
        Me.LinkLblCAPSignup.TabIndex = 9
        Me.LinkLblCAPSignup.TabStop = True
        Me.LinkLblCAPSignup.Text = "CAP Signup/Maintenance"
        '
        'lblCAPParticipationLevel
        '
        Me.lblCAPParticipationLevel.Location = New System.Drawing.Point(680, 48)
        Me.lblCAPParticipationLevel.Name = "lblCAPParticipationLevel"
        Me.lblCAPParticipationLevel.Size = New System.Drawing.Size(264, 20)
        Me.lblCAPParticipationLevel.TabIndex = 1004
        Me.lblCAPParticipationLevel.Text = "NONE - 0/0 (Compliant/Candidate)"
        Me.lblCAPParticipationLevel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlOwnerName
        '
        Me.pnlOwnerName.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.pnlOwnerName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlOwnerName.Controls.Add(Me.pnlOwnerNameButton)
        Me.pnlOwnerName.Controls.Add(Me.pnlOwnerOrg)
        Me.pnlOwnerName.Controls.Add(Me.pnlPersonOrganization)
        Me.pnlOwnerName.Controls.Add(Me.pnlOwnerPerson)
        Me.pnlOwnerName.Location = New System.Drawing.Point(616, 232)
        Me.pnlOwnerName.Name = "pnlOwnerName"
        Me.pnlOwnerName.Size = New System.Drawing.Size(296, 256)
        Me.pnlOwnerName.TabIndex = 2
        Me.pnlOwnerName.Visible = False
        '
        'pnlOwnerNameButton
        '
        Me.pnlOwnerNameButton.Controls.Add(Me.btnOwnerNameClose)
        Me.pnlOwnerNameButton.Controls.Add(Me.btnOwnerNameCancel)
        Me.pnlOwnerNameButton.Controls.Add(Me.btnOwnerNameOK)
        Me.pnlOwnerNameButton.Location = New System.Drawing.Point(0, 224)
        Me.pnlOwnerNameButton.Name = "pnlOwnerNameButton"
        Me.pnlOwnerNameButton.Size = New System.Drawing.Size(304, 25)
        Me.pnlOwnerNameButton.TabIndex = 3
        '
        'btnOwnerNameClose
        '
        Me.btnOwnerNameClose.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerNameClose.Location = New System.Drawing.Point(168, 0)
        Me.btnOwnerNameClose.Name = "btnOwnerNameClose"
        Me.btnOwnerNameClose.Size = New System.Drawing.Size(60, 24)
        Me.btnOwnerNameClose.TabIndex = 2
        Me.btnOwnerNameClose.Text = "Close"
        '
        'btnOwnerNameCancel
        '
        Me.btnOwnerNameCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerNameCancel.Enabled = False
        Me.btnOwnerNameCancel.Location = New System.Drawing.Point(104, 0)
        Me.btnOwnerNameCancel.Name = "btnOwnerNameCancel"
        Me.btnOwnerNameCancel.Size = New System.Drawing.Size(60, 24)
        Me.btnOwnerNameCancel.TabIndex = 1
        Me.btnOwnerNameCancel.Text = "Cancel"
        '
        'btnOwnerNameOK
        '
        Me.btnOwnerNameOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerNameOK.Enabled = False
        Me.btnOwnerNameOK.Location = New System.Drawing.Point(56, 1)
        Me.btnOwnerNameOK.Name = "btnOwnerNameOK"
        Me.btnOwnerNameOK.Size = New System.Drawing.Size(44, 24)
        Me.btnOwnerNameOK.TabIndex = 0
        Me.btnOwnerNameOK.Text = "Save"
        '
        'pnlOwnerOrg
        '
        Me.pnlOwnerOrg.Controls.Add(Me.txtOwnerOrgName)
        Me.pnlOwnerOrg.Controls.Add(Me.lblOwnerOrgName)
        Me.pnlOwnerOrg.Controls.Add(Me.cmbOwnerOrgEntityCode)
        Me.pnlOwnerOrg.Controls.Add(Me.lblOwnerOrgEntityCode)
        Me.pnlOwnerOrg.Location = New System.Drawing.Point(0, 160)
        Me.pnlOwnerOrg.Name = "pnlOwnerOrg"
        Me.pnlOwnerOrg.Size = New System.Drawing.Size(304, 56)
        Me.pnlOwnerOrg.TabIndex = 2
        '
        'txtOwnerOrgName
        '
        Me.txtOwnerOrgName.Location = New System.Drawing.Point(97, 7)
        Me.txtOwnerOrgName.Name = "txtOwnerOrgName"
        Me.txtOwnerOrgName.Size = New System.Drawing.Size(192, 21)
        Me.txtOwnerOrgName.TabIndex = 0
        Me.txtOwnerOrgName.Tag = ""
        Me.txtOwnerOrgName.Text = ""
        '
        'lblOwnerOrgName
        '
        Me.lblOwnerOrgName.Location = New System.Drawing.Point(8, 7)
        Me.lblOwnerOrgName.Name = "lblOwnerOrgName"
        Me.lblOwnerOrgName.TabIndex = 88
        Me.lblOwnerOrgName.Text = "Name"
        '
        'cmbOwnerOrgEntityCode
        '
        Me.cmbOwnerOrgEntityCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerOrgEntityCode.DropDownWidth = 200
        Me.cmbOwnerOrgEntityCode.ItemHeight = 15
        Me.cmbOwnerOrgEntityCode.Items.AddRange(New Object() {"Owner Type 1", "Owner Type 2"})
        Me.cmbOwnerOrgEntityCode.Location = New System.Drawing.Point(96, 32)
        Me.cmbOwnerOrgEntityCode.Name = "cmbOwnerOrgEntityCode"
        Me.cmbOwnerOrgEntityCode.Size = New System.Drawing.Size(152, 23)
        Me.cmbOwnerOrgEntityCode.TabIndex = 1
        '
        'lblOwnerOrgEntityCode
        '
        Me.lblOwnerOrgEntityCode.Location = New System.Drawing.Point(8, 32)
        Me.lblOwnerOrgEntityCode.Name = "lblOwnerOrgEntityCode"
        Me.lblOwnerOrgEntityCode.Size = New System.Drawing.Size(72, 23)
        Me.lblOwnerOrgEntityCode.TabIndex = 91
        Me.lblOwnerOrgEntityCode.Text = "Entity Code"
        '
        'pnlPersonOrganization
        '
        Me.pnlPersonOrganization.Controls.Add(Me.btnOwnerNameSearch)
        Me.pnlPersonOrganization.Controls.Add(Me.rdOwnerOrg)
        Me.pnlPersonOrganization.Controls.Add(Me.rdOwnerPerson)
        Me.pnlPersonOrganization.Location = New System.Drawing.Point(0, 0)
        Me.pnlPersonOrganization.Name = "pnlPersonOrganization"
        Me.pnlPersonOrganization.Size = New System.Drawing.Size(280, 24)
        Me.pnlPersonOrganization.TabIndex = 0
        '
        'btnOwnerNameSearch
        '
        Me.btnOwnerNameSearch.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerNameSearch.Enabled = False
        Me.btnOwnerNameSearch.Location = New System.Drawing.Point(200, 2)
        Me.btnOwnerNameSearch.Name = "btnOwnerNameSearch"
        Me.btnOwnerNameSearch.Size = New System.Drawing.Size(16, 20)
        Me.btnOwnerNameSearch.TabIndex = 2
        Me.btnOwnerNameSearch.Text = "?"
        '
        'rdOwnerOrg
        '
        Me.rdOwnerOrg.Location = New System.Drawing.Point(88, 3)
        Me.rdOwnerOrg.Name = "rdOwnerOrg"
        Me.rdOwnerOrg.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.rdOwnerOrg.Size = New System.Drawing.Size(93, 24)
        Me.rdOwnerOrg.TabIndex = 1
        Me.rdOwnerOrg.Text = "Organization"
        Me.rdOwnerOrg.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'rdOwnerPerson
        '
        Me.rdOwnerPerson.Checked = True
        Me.rdOwnerPerson.Location = New System.Drawing.Point(8, 3)
        Me.rdOwnerPerson.Name = "rdOwnerPerson"
        Me.rdOwnerPerson.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.rdOwnerPerson.Size = New System.Drawing.Size(64, 24)
        Me.rdOwnerPerson.TabIndex = 0
        Me.rdOwnerPerson.TabStop = True
        Me.rdOwnerPerson.Text = "Person"
        Me.rdOwnerPerson.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'pnlOwnerPerson
        '
        Me.pnlOwnerPerson.Controls.Add(Me.cmbOwnerNameSuffix)
        Me.pnlOwnerPerson.Controls.Add(Me.cmbOwnerNameTitle)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerNameSuffix)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerNameTitle)
        Me.pnlOwnerPerson.Controls.Add(Me.txtOwnerMiddleName)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerMiddleName)
        Me.pnlOwnerPerson.Controls.Add(Me.txtOwnerLastName)
        Me.pnlOwnerPerson.Controls.Add(Me.txtOwnerFirstName)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerLastName)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerFirstName)
        Me.pnlOwnerPerson.Location = New System.Drawing.Point(0, 29)
        Me.pnlOwnerPerson.Name = "pnlOwnerPerson"
        Me.pnlOwnerPerson.Size = New System.Drawing.Size(296, 131)
        Me.pnlOwnerPerson.TabIndex = 1
        '
        'cmbOwnerNameSuffix
        '
        Me.cmbOwnerNameSuffix.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerNameSuffix.Items.AddRange(New Object() {"Jr", "Sr", "I", "II", "III", "IV", "V", "VI"})
        Me.cmbOwnerNameSuffix.Location = New System.Drawing.Point(96, 103)
        Me.cmbOwnerNameSuffix.Name = "cmbOwnerNameSuffix"
        Me.cmbOwnerNameSuffix.Size = New System.Drawing.Size(80, 23)
        Me.cmbOwnerNameSuffix.TabIndex = 4
        '
        'cmbOwnerNameTitle
        '
        Me.cmbOwnerNameTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerNameTitle.Items.AddRange(New Object() {"Mr", "Mrs", "Ms", "Dr", "Sir"})
        Me.cmbOwnerNameTitle.Location = New System.Drawing.Point(96, 8)
        Me.cmbOwnerNameTitle.Name = "cmbOwnerNameTitle"
        Me.cmbOwnerNameTitle.Size = New System.Drawing.Size(80, 23)
        Me.cmbOwnerNameTitle.TabIndex = 0
        '
        'lblOwnerNameSuffix
        '
        Me.lblOwnerNameSuffix.Location = New System.Drawing.Point(8, 103)
        Me.lblOwnerNameSuffix.Name = "lblOwnerNameSuffix"
        Me.lblOwnerNameSuffix.TabIndex = 92
        Me.lblOwnerNameSuffix.Text = "Suffix"
        '
        'lblOwnerNameTitle
        '
        Me.lblOwnerNameTitle.Location = New System.Drawing.Point(8, 8)
        Me.lblOwnerNameTitle.Name = "lblOwnerNameTitle"
        Me.lblOwnerNameTitle.TabIndex = 91
        Me.lblOwnerNameTitle.Text = "Title"
        '
        'txtOwnerMiddleName
        '
        Me.txtOwnerMiddleName.Location = New System.Drawing.Point(96, 56)
        Me.txtOwnerMiddleName.Name = "txtOwnerMiddleName"
        Me.txtOwnerMiddleName.Size = New System.Drawing.Size(192, 21)
        Me.txtOwnerMiddleName.TabIndex = 2
        Me.txtOwnerMiddleName.Tag = ""
        Me.txtOwnerMiddleName.Text = ""
        '
        'lblOwnerMiddleName
        '
        Me.lblOwnerMiddleName.Location = New System.Drawing.Point(8, 56)
        Me.lblOwnerMiddleName.Name = "lblOwnerMiddleName"
        Me.lblOwnerMiddleName.Size = New System.Drawing.Size(100, 17)
        Me.lblOwnerMiddleName.TabIndex = 89
        Me.lblOwnerMiddleName.Text = "Middle Name"
        '
        'txtOwnerLastName
        '
        Me.txtOwnerLastName.Location = New System.Drawing.Point(96, 80)
        Me.txtOwnerLastName.Name = "txtOwnerLastName"
        Me.txtOwnerLastName.Size = New System.Drawing.Size(192, 21)
        Me.txtOwnerLastName.TabIndex = 3
        Me.txtOwnerLastName.Tag = ""
        Me.txtOwnerLastName.Text = ""
        '
        'txtOwnerFirstName
        '
        Me.txtOwnerFirstName.Location = New System.Drawing.Point(96, 32)
        Me.txtOwnerFirstName.Name = "txtOwnerFirstName"
        Me.txtOwnerFirstName.Size = New System.Drawing.Size(192, 21)
        Me.txtOwnerFirstName.TabIndex = 1
        Me.txtOwnerFirstName.Tag = ""
        Me.txtOwnerFirstName.Text = ""
        '
        'lblOwnerLastName
        '
        Me.lblOwnerLastName.Location = New System.Drawing.Point(8, 80)
        Me.lblOwnerLastName.Name = "lblOwnerLastName"
        Me.lblOwnerLastName.TabIndex = 86
        Me.lblOwnerLastName.Text = "Last Name"
        '
        'lblOwnerFirstName
        '
        Me.lblOwnerFirstName.Location = New System.Drawing.Point(8, 32)
        Me.lblOwnerFirstName.Name = "lblOwnerFirstName"
        Me.lblOwnerFirstName.TabIndex = 85
        Me.lblOwnerFirstName.Text = "First Name"
        '
        'mskTxtOwnerFax
        '
        Me.mskTxtOwnerFax.ContainingControl = Me
        Me.mskTxtOwnerFax.Location = New System.Drawing.Point(424, 112)
        Me.mskTxtOwnerFax.Name = "mskTxtOwnerFax"
        Me.mskTxtOwnerFax.OcxState = CType(resources.GetObject("mskTxtOwnerFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerFax.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerFax.TabIndex = 7
        '
        'mskTxtOwnerPhone2
        '
        Me.mskTxtOwnerPhone2.ContainingControl = Me
        Me.mskTxtOwnerPhone2.Location = New System.Drawing.Point(424, 88)
        Me.mskTxtOwnerPhone2.Name = "mskTxtOwnerPhone2"
        Me.mskTxtOwnerPhone2.OcxState = CType(resources.GetObject("mskTxtOwnerPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone2.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerPhone2.TabIndex = 6
        '
        'mskTxtOwnerPhone
        '
        Me.mskTxtOwnerPhone.ContainingControl = Me
        Me.mskTxtOwnerPhone.Location = New System.Drawing.Point(424, 64)
        Me.mskTxtOwnerPhone.Name = "mskTxtOwnerPhone"
        Me.mskTxtOwnerPhone.OcxState = CType(resources.GetObject("mskTxtOwnerPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerPhone.TabIndex = 5
        '
        'lblOwnerEmail
        '
        Me.lblOwnerEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerEmail.Location = New System.Drawing.Point(552, 104)
        Me.lblOwnerEmail.Name = "lblOwnerEmail"
        Me.lblOwnerEmail.Size = New System.Drawing.Size(40, 20)
        Me.lblOwnerEmail.TabIndex = 11
        Me.lblOwnerEmail.Text = "Email:"
        Me.lblOwnerEmail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOwnerEmail
        '
        Me.txtOwnerEmail.AcceptsTab = True
        Me.txtOwnerEmail.AutoSize = False
        Me.txtOwnerEmail.Location = New System.Drawing.Point(592, 104)
        Me.txtOwnerEmail.Name = "txtOwnerEmail"
        Me.txtOwnerEmail.Size = New System.Drawing.Size(200, 20)
        Me.txtOwnerEmail.TabIndex = 10
        Me.txtOwnerEmail.Text = ""
        Me.txtOwnerEmail.WordWrap = False
        '
        'lblFax
        '
        Me.lblFax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFax.Location = New System.Drawing.Point(344, 112)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(72, 20)
        Me.lblFax.TabIndex = 44
        Me.lblFax.Text = "Fax:"
        '
        'pnlOwnerButtons
        '
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerFlag)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerComment)
        Me.pnlOwnerButtons.Controls.Add(Me.btnTransferOwnership)
        Me.pnlOwnerButtons.Controls.Add(Me.btnSaveOwner)
        Me.pnlOwnerButtons.Controls.Add(Me.btnDeleteOwner)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerCancel)
        Me.pnlOwnerButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.pnlOwnerButtons.Location = New System.Drawing.Point(336, 183)
        Me.pnlOwnerButtons.Name = "pnlOwnerButtons"
        Me.pnlOwnerButtons.Size = New System.Drawing.Size(543, 40)
        Me.pnlOwnerButtons.TabIndex = 11
        '
        'btnOwnerFlag
        '
        Me.btnOwnerFlag.Location = New System.Drawing.Point(157, 8)
        Me.btnOwnerFlag.Name = "btnOwnerFlag"
        Me.btnOwnerFlag.Size = New System.Drawing.Size(75, 26)
        Me.btnOwnerFlag.TabIndex = 4
        Me.btnOwnerFlag.Text = "Flags"
        '
        'btnOwnerComment
        '
        Me.btnOwnerComment.Location = New System.Drawing.Point(240, 8)
        Me.btnOwnerComment.Name = "btnOwnerComment"
        Me.btnOwnerComment.Size = New System.Drawing.Size(80, 26)
        Me.btnOwnerComment.TabIndex = 5
        Me.btnOwnerComment.Text = "Comments"
        '
        'btnTransferOwnership
        '
        Me.btnTransferOwnership.BackColor = System.Drawing.SystemColors.Control
        Me.btnTransferOwnership.Enabled = False
        Me.btnTransferOwnership.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTransferOwnership.Location = New System.Drawing.Point(424, 8)
        Me.btnTransferOwnership.Name = "btnTransferOwnership"
        Me.btnTransferOwnership.Size = New System.Drawing.Size(112, 26)
        Me.btnTransferOwnership.TabIndex = 3
        Me.btnTransferOwnership.Text = "Transfer Ownership"
        '
        'btnSaveOwner
        '
        Me.btnSaveOwner.BackColor = System.Drawing.SystemColors.Control
        Me.btnSaveOwner.Enabled = False
        Me.btnSaveOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveOwner.Location = New System.Drawing.Point(8, 8)
        Me.btnSaveOwner.Name = "btnSaveOwner"
        Me.btnSaveOwner.Size = New System.Drawing.Size(88, 26)
        Me.btnSaveOwner.TabIndex = 0
        Me.btnSaveOwner.Text = "Save Owner"
        '
        'btnDeleteOwner
        '
        Me.btnDeleteOwner.BackColor = System.Drawing.SystemColors.Control
        Me.btnDeleteOwner.Enabled = False
        Me.btnDeleteOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteOwner.Location = New System.Drawing.Point(328, 8)
        Me.btnDeleteOwner.Name = "btnDeleteOwner"
        Me.btnDeleteOwner.Size = New System.Drawing.Size(88, 26)
        Me.btnDeleteOwner.TabIndex = 2
        Me.btnDeleteOwner.Text = "Delete Owner"
        '
        'btnOwnerCancel
        '
        Me.btnOwnerCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerCancel.Enabled = False
        Me.btnOwnerCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerCancel.Location = New System.Drawing.Point(101, 8)
        Me.btnOwnerCancel.Name = "btnOwnerCancel"
        Me.btnOwnerCancel.Size = New System.Drawing.Size(51, 26)
        Me.btnOwnerCancel.TabIndex = 1
        Me.btnOwnerCancel.Text = "Cancel"
        '
        'txtOwnerAddress
        '
        Me.txtOwnerAddress.Location = New System.Drawing.Point(80, 56)
        Me.txtOwnerAddress.Multiline = True
        Me.txtOwnerAddress.Name = "txtOwnerAddress"
        Me.txtOwnerAddress.ReadOnly = True
        Me.txtOwnerAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtOwnerAddress.Size = New System.Drawing.Size(248, 103)
        Me.txtOwnerAddress.TabIndex = 2
        Me.txtOwnerAddress.Text = ""
        Me.txtOwnerAddress.WordWrap = False
        '
        'lblOwnerAddress
        '
        Me.lblOwnerAddress.Location = New System.Drawing.Point(7, 56)
        Me.lblOwnerAddress.Name = "lblOwnerAddress"
        Me.lblOwnerAddress.Size = New System.Drawing.Size(72, 16)
        Me.lblOwnerAddress.TabIndex = 88
        Me.lblOwnerAddress.Text = "Address:"
        '
        'txtOwnerName
        '
        Me.txtOwnerName.Location = New System.Drawing.Point(80, 32)
        Me.txtOwnerName.Name = "txtOwnerName"
        Me.txtOwnerName.ReadOnly = True
        Me.txtOwnerName.Size = New System.Drawing.Size(248, 21)
        Me.txtOwnerName.TabIndex = 1
        Me.txtOwnerName.Text = ""
        '
        'lblOwnerName
        '
        Me.lblOwnerName.Location = New System.Drawing.Point(8, 32)
        Me.lblOwnerName.Name = "lblOwnerName"
        Me.lblOwnerName.Size = New System.Drawing.Size(70, 20)
        Me.lblOwnerName.TabIndex = 86
        Me.lblOwnerName.Text = "Name:"
        '
        'lblOwnerStatus
        '
        Me.lblOwnerStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerStatus.Location = New System.Drawing.Point(344, 16)
        Me.lblOwnerStatus.Name = "lblOwnerStatus"
        Me.lblOwnerStatus.Size = New System.Drawing.Size(78, 20)
        Me.lblOwnerStatus.TabIndex = 84
        Me.lblOwnerStatus.Text = "Owner Status:"
        '
        'lblOwnerCapParticipant
        '
        Me.lblOwnerCapParticipant.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerCapParticipant.Location = New System.Drawing.Point(552, 48)
        Me.lblOwnerCapParticipant.Name = "lblOwnerCapParticipant"
        Me.lblOwnerCapParticipant.Size = New System.Drawing.Size(128, 20)
        Me.lblOwnerCapParticipant.TabIndex = 52
        Me.lblOwnerCapParticipant.Text = "CAP Participation Level"
        Me.lblOwnerCapParticipant.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPhone2
        '
        Me.lblPhone2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhone2.Location = New System.Drawing.Point(344, 88)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(70, 20)
        Me.lblPhone2.TabIndex = 45
        Me.lblPhone2.Text = "Phone 2:"
        '
        'lblOwnerType
        '
        Me.lblOwnerType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerType.Location = New System.Drawing.Point(8, 160)
        Me.lblOwnerType.Name = "lblOwnerType"
        Me.lblOwnerType.Size = New System.Drawing.Size(72, 23)
        Me.lblOwnerType.TabIndex = 40
        Me.lblOwnerType.Text = "Owner Type:"
        '
        'txtOwnerAIID
        '
        Me.txtOwnerAIID.AcceptsTab = True
        Me.txtOwnerAIID.AutoSize = False
        Me.txtOwnerAIID.Enabled = False
        Me.txtOwnerAIID.Location = New System.Drawing.Point(424, 40)
        Me.txtOwnerAIID.Name = "txtOwnerAIID"
        Me.txtOwnerAIID.Size = New System.Drawing.Size(96, 20)
        Me.txtOwnerAIID.TabIndex = 4
        Me.txtOwnerAIID.Text = ""
        Me.txtOwnerAIID.WordWrap = False
        '
        'lblOwnerAIID
        '
        Me.lblOwnerAIID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerAIID.Location = New System.Drawing.Point(344, 40)
        Me.lblOwnerAIID.Name = "lblOwnerAIID"
        Me.lblOwnerAIID.Size = New System.Drawing.Size(70, 20)
        Me.lblOwnerAIID.TabIndex = 38
        Me.lblOwnerAIID.Text = "Ensite ID:"
        '
        'lblOwnerIDValue
        '
        Me.lblOwnerIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerIDValue.Location = New System.Drawing.Point(86, 8)
        Me.lblOwnerIDValue.Name = "lblOwnerIDValue"
        Me.lblOwnerIDValue.Size = New System.Drawing.Size(96, 23)
        Me.lblOwnerIDValue.TabIndex = 0
        '
        'lblOwnerID
        '
        Me.lblOwnerID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerID.Location = New System.Drawing.Point(8, 8)
        Me.lblOwnerID.Name = "lblOwnerID"
        Me.lblOwnerID.Size = New System.Drawing.Size(64, 20)
        Me.lblOwnerID.TabIndex = 36
        Me.lblOwnerID.Text = "Owner ID:"
        '
        'lblOwnerPhone
        '
        Me.lblOwnerPhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerPhone.Location = New System.Drawing.Point(344, 64)
        Me.lblOwnerPhone.Name = "lblOwnerPhone"
        Me.lblOwnerPhone.Size = New System.Drawing.Size(70, 20)
        Me.lblOwnerPhone.TabIndex = 32
        Me.lblOwnerPhone.Text = "Phone:"
        '
        'cmbOwnerType
        '
        Me.cmbOwnerType.DisplayMember = "1"
        Me.cmbOwnerType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerType.DropDownWidth = 200
        Me.cmbOwnerType.ItemHeight = 15
        Me.cmbOwnerType.Location = New System.Drawing.Point(80, 160)
        Me.cmbOwnerType.Name = "cmbOwnerType"
        Me.cmbOwnerType.Size = New System.Drawing.Size(248, 23)
        Me.cmbOwnerType.TabIndex = 3
        Me.cmbOwnerType.ValueMember = "1"
        '
        'chkCAPParticipant
        '
        Me.chkCAPParticipant.Checked = True
        Me.chkCAPParticipant.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCAPParticipant.Location = New System.Drawing.Point(808, 16)
        Me.chkCAPParticipant.Name = "chkCAPParticipant"
        Me.chkCAPParticipant.Size = New System.Drawing.Size(16, 20)
        Me.chkCAPParticipant.TabIndex = 28
        Me.chkCAPParticipant.Visible = False
        '
        'tbPageFacilityDetail
        '
        Me.tbPageFacilityDetail.Controls.Add(Me.pnlFacilityBottom)
        Me.tbPageFacilityDetail.Controls.Add(Me.pnl_FacilityDetail)
        Me.tbPageFacilityDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageFacilityDetail.Name = "tbPageFacilityDetail"
        Me.tbPageFacilityDetail.Size = New System.Drawing.Size(964, 636)
        Me.tbPageFacilityDetail.TabIndex = 8
        Me.tbPageFacilityDetail.Text = "Facility Details"
        '
        'pnlFacilityBottom
        '
        Me.pnlFacilityBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlFacilityBottom.Controls.Add(Me.tabCtrlFacilityTankPipe)
        Me.pnlFacilityBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityBottom.Location = New System.Drawing.Point(0, 280)
        Me.pnlFacilityBottom.Name = "pnlFacilityBottom"
        Me.pnlFacilityBottom.Size = New System.Drawing.Size(964, 356)
        Me.pnlFacilityBottom.TabIndex = 3
        '
        'tabCtrlFacilityTankPipe
        '
        Me.tabCtrlFacilityTankPipe.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.tabCtrlFacilityTankPipe.Controls.Add(Me.tbPageTankPipe)
        Me.tabCtrlFacilityTankPipe.Controls.Add(Me.tbPageFacilityContactList)
        Me.tabCtrlFacilityTankPipe.Controls.Add(Me.tpPreviouslyOwnedOwners)
        Me.tabCtrlFacilityTankPipe.Controls.Add(Me.tbPageFacilityDocuments)
        Me.tabCtrlFacilityTankPipe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabCtrlFacilityTankPipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabCtrlFacilityTankPipe.ItemSize = New System.Drawing.Size(107, 23)
        Me.tabCtrlFacilityTankPipe.Location = New System.Drawing.Point(0, 0)
        Me.tabCtrlFacilityTankPipe.Name = "tabCtrlFacilityTankPipe"
        Me.tabCtrlFacilityTankPipe.SelectedIndex = 0
        Me.tabCtrlFacilityTankPipe.Size = New System.Drawing.Size(962, 354)
        Me.tabCtrlFacilityTankPipe.TabIndex = 2
        '
        'tbPageTankPipe
        '
        Me.tbPageTankPipe.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageTankPipe.Controls.Add(Me.pnlTankPipe)
        Me.tbPageTankPipe.Controls.Add(Me.Label4)
        Me.tbPageTankPipe.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPageTankPipe.Location = New System.Drawing.Point(4, 27)
        Me.tbPageTankPipe.Name = "tbPageTankPipe"
        Me.tbPageTankPipe.Size = New System.Drawing.Size(954, 323)
        Me.tbPageTankPipe.TabIndex = 0
        Me.tbPageTankPipe.Text = "Tank-Pipe"
        '
        'pnlTankPipe
        '
        Me.pnlTankPipe.AutoScroll = True
        Me.pnlTankPipe.Controls.Add(Me.btnExpand)
        Me.pnlTankPipe.Controls.Add(Me.btnAddTank)
        Me.pnlTankPipe.Controls.Add(Me.dgPipesAndTanks)
        Me.pnlTankPipe.Controls.Add(Me.Label12)
        Me.pnlTankPipe.Controls.Add(Me.pnlFacilityTankPipeButton)
        Me.pnlTankPipe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlTankPipe.Location = New System.Drawing.Point(0, 0)
        Me.pnlTankPipe.Name = "pnlTankPipe"
        Me.pnlTankPipe.Size = New System.Drawing.Size(950, 319)
        Me.pnlTankPipe.TabIndex = 95
        '
        'btnExpand
        '
        Me.btnExpand.Location = New System.Drawing.Point(97, 0)
        Me.btnExpand.Name = "btnExpand"
        Me.btnExpand.Size = New System.Drawing.Size(96, 22)
        Me.btnExpand.TabIndex = 145
        Me.btnExpand.Text = "Expand All"
        Me.btnExpand.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnAddTank
        '
        Me.btnAddTank.Location = New System.Drawing.Point(0, 0)
        Me.btnAddTank.Name = "btnAddTank"
        Me.btnAddTank.Size = New System.Drawing.Size(96, 22)
        Me.btnAddTank.TabIndex = 142
        Me.btnAddTank.Text = "Add Tank"
        Me.btnAddTank.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dgPipesAndTanks
        '
        Me.dgPipesAndTanks.Cursor = System.Windows.Forms.Cursors.Default
        Me.dgPipesAndTanks.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgPipesAndTanks.Location = New System.Drawing.Point(0, 0)
        Me.dgPipesAndTanks.Name = "dgPipesAndTanks"
        Me.dgPipesAndTanks.Size = New System.Drawing.Size(950, 295)
        Me.dgPipesAndTanks.TabIndex = 10
        Me.dgPipesAndTanks.Text = "Tanks And Pipes"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(822, 136)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(32, 8)
        Me.Label12.TabIndex = 9
        '
        'pnlFacilityTankPipeButton
        '
        Me.pnlFacilityTankPipeButton.Controls.Add(Me.lblTotalNoOfTanksValue)
        Me.pnlFacilityTankPipeButton.Controls.Add(Me.lblTotalNoOfTanks)
        Me.pnlFacilityTankPipeButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFacilityTankPipeButton.Location = New System.Drawing.Point(0, 295)
        Me.pnlFacilityTankPipeButton.Name = "pnlFacilityTankPipeButton"
        Me.pnlFacilityTankPipeButton.Size = New System.Drawing.Size(950, 24)
        Me.pnlFacilityTankPipeButton.TabIndex = 96
        '
        'lblTotalNoOfTanksValue
        '
        Me.lblTotalNoOfTanksValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfTanksValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfTanksValue.Location = New System.Drawing.Point(200, 0)
        Me.lblTotalNoOfTanksValue.Name = "lblTotalNoOfTanksValue"
        Me.lblTotalNoOfTanksValue.Size = New System.Drawing.Size(48, 24)
        Me.lblTotalNoOfTanksValue.TabIndex = 5
        '
        'lblTotalNoOfTanks
        '
        Me.lblTotalNoOfTanks.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfTanks.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfTanks.Location = New System.Drawing.Point(0, 0)
        Me.lblTotalNoOfTanks.Name = "lblTotalNoOfTanks"
        Me.lblTotalNoOfTanks.Size = New System.Drawing.Size(200, 24)
        Me.lblTotalNoOfTanks.TabIndex = 4
        Me.lblTotalNoOfTanks.Text = "Number of Tanks at this Location:"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(1008, 104)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(7, 23)
        Me.Label4.TabIndex = 9
        '
        'tbPageFacilityContactList
        '
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactContainer)
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactHeader)
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactBottom)
        Me.tbPageFacilityContactList.Location = New System.Drawing.Point(4, 27)
        Me.tbPageFacilityContactList.Name = "tbPageFacilityContactList"
        Me.tbPageFacilityContactList.Size = New System.Drawing.Size(954, 323)
        Me.tbPageFacilityContactList.TabIndex = 4
        Me.tbPageFacilityContactList.Text = "Contacts"
        '
        'pnlFacilityContactContainer
        '
        Me.pnlFacilityContactContainer.Controls.Add(Me.ugFacilityContacts)
        Me.pnlFacilityContactContainer.Controls.Add(Me.Label2)
        Me.pnlFacilityContactContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityContactContainer.Location = New System.Drawing.Point(0, 25)
        Me.pnlFacilityContactContainer.Name = "pnlFacilityContactContainer"
        Me.pnlFacilityContactContainer.Size = New System.Drawing.Size(954, 268)
        Me.pnlFacilityContactContainer.TabIndex = 2
        '
        'ugFacilityContacts
        '
        Me.ugFacilityContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilityContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFacilityContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFacilityContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugFacilityContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFacilityContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugFacilityContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFacilityContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugFacilityContacts.Name = "ugFacilityContacts"
        Me.ugFacilityContacts.Size = New System.Drawing.Size(954, 268)
        Me.ugFacilityContacts.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(792, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(7, 23)
        Me.Label2.TabIndex = 2
        '
        'pnlFacilityContactHeader
        '
        Me.pnlFacilityContactHeader.Controls.Add(Me.chkFacilityShowActiveContactOnly)
        Me.pnlFacilityContactHeader.Controls.Add(Me.chkFacilityShowRelatedContacts)
        Me.pnlFacilityContactHeader.Controls.Add(Me.chkFacilityShowContactsforAllModule)
        Me.pnlFacilityContactHeader.Controls.Add(Me.lblFacilityContacts)
        Me.pnlFacilityContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFacilityContactHeader.DockPadding.All = 3
        Me.pnlFacilityContactHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlFacilityContactHeader.Name = "pnlFacilityContactHeader"
        Me.pnlFacilityContactHeader.Size = New System.Drawing.Size(954, 25)
        Me.pnlFacilityContactHeader.TabIndex = 1
        '
        'chkFacilityShowActiveContactOnly
        '
        Me.chkFacilityShowActiveContactOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFacilityShowActiveContactOnly.Location = New System.Drawing.Point(635, 7)
        Me.chkFacilityShowActiveContactOnly.Name = "chkFacilityShowActiveContactOnly"
        Me.chkFacilityShowActiveContactOnly.Size = New System.Drawing.Size(144, 16)
        Me.chkFacilityShowActiveContactOnly.TabIndex = 2
        Me.chkFacilityShowActiveContactOnly.Tag = "646"
        Me.chkFacilityShowActiveContactOnly.Text = "Show Active Only"
        '
        'chkFacilityShowRelatedContacts
        '
        Me.chkFacilityShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFacilityShowRelatedContacts.Location = New System.Drawing.Point(467, 7)
        Me.chkFacilityShowRelatedContacts.Name = "chkFacilityShowRelatedContacts"
        Me.chkFacilityShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkFacilityShowRelatedContacts.TabIndex = 1
        Me.chkFacilityShowRelatedContacts.Tag = "645"
        Me.chkFacilityShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkFacilityShowContactsforAllModule
        '
        Me.chkFacilityShowContactsforAllModule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFacilityShowContactsforAllModule.Location = New System.Drawing.Point(251, 7)
        Me.chkFacilityShowContactsforAllModule.Name = "chkFacilityShowContactsforAllModule"
        Me.chkFacilityShowContactsforAllModule.Size = New System.Drawing.Size(200, 16)
        Me.chkFacilityShowContactsforAllModule.TabIndex = 0
        Me.chkFacilityShowContactsforAllModule.Tag = "644"
        Me.chkFacilityShowContactsforAllModule.Text = "Show Contacts for All Modules"
        '
        'lblFacilityContacts
        '
        Me.lblFacilityContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityContacts.Location = New System.Drawing.Point(8, 7)
        Me.lblFacilityContacts.Name = "lblFacilityContacts"
        Me.lblFacilityContacts.Size = New System.Drawing.Size(104, 16)
        Me.lblFacilityContacts.TabIndex = 139
        Me.lblFacilityContacts.Text = "Facility Contacts"
        '
        'pnlFacilityContactBottom
        '
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityModifyContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityDeleteContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityAssociateContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityAddSearchContact)
        Me.pnlFacilityContactBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFacilityContactBottom.DockPadding.All = 3
        Me.pnlFacilityContactBottom.Location = New System.Drawing.Point(0, 293)
        Me.pnlFacilityContactBottom.Name = "pnlFacilityContactBottom"
        Me.pnlFacilityContactBottom.Size = New System.Drawing.Size(954, 30)
        Me.pnlFacilityContactBottom.TabIndex = 2
        '
        'btnFacilityModifyContact
        '
        Me.btnFacilityModifyContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityModifyContact.Location = New System.Drawing.Point(273, 4)
        Me.btnFacilityModifyContact.Name = "btnFacilityModifyContact"
        Me.btnFacilityModifyContact.Size = New System.Drawing.Size(120, 23)
        Me.btnFacilityModifyContact.TabIndex = 1
        Me.btnFacilityModifyContact.Text = "Modify Contact"
        '
        'btnFacilityDeleteContact
        '
        Me.btnFacilityDeleteContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityDeleteContact.Location = New System.Drawing.Point(401, 4)
        Me.btnFacilityDeleteContact.Name = "btnFacilityDeleteContact"
        Me.btnFacilityDeleteContact.Size = New System.Drawing.Size(112, 23)
        Me.btnFacilityDeleteContact.TabIndex = 2
        Me.btnFacilityDeleteContact.Text = "Delete Contact"
        '
        'btnFacilityAssociateContact
        '
        Me.btnFacilityAssociateContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityAssociateContact.Location = New System.Drawing.Point(521, 4)
        Me.btnFacilityAssociateContact.Name = "btnFacilityAssociateContact"
        Me.btnFacilityAssociateContact.Size = New System.Drawing.Size(128, 23)
        Me.btnFacilityAssociateContact.TabIndex = 3
        Me.btnFacilityAssociateContact.Text = "Associate Contact"
        '
        'btnFacilityAddSearchContact
        '
        Me.btnFacilityAddSearchContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityAddSearchContact.Location = New System.Drawing.Point(129, 4)
        Me.btnFacilityAddSearchContact.Name = "btnFacilityAddSearchContact"
        Me.btnFacilityAddSearchContact.Size = New System.Drawing.Size(136, 23)
        Me.btnFacilityAddSearchContact.TabIndex = 0
        Me.btnFacilityAddSearchContact.Text = "Add/Search Contact"
        '
        'tpPreviouslyOwnedOwners
        '
        Me.tpPreviouslyOwnedOwners.Controls.Add(Me.ugPreviouslyOwnedOwners)
        Me.tpPreviouslyOwnedOwners.Controls.Add(Me.Panel3)
        Me.tpPreviouslyOwnedOwners.Location = New System.Drawing.Point(4, 27)
        Me.tpPreviouslyOwnedOwners.Name = "tpPreviouslyOwnedOwners"
        Me.tpPreviouslyOwnedOwners.Size = New System.Drawing.Size(954, 323)
        Me.tpPreviouslyOwnedOwners.TabIndex = 3
        Me.tpPreviouslyOwnedOwners.Text = "Previous Owners"
        '
        'ugPreviouslyOwnedOwners
        '
        Me.ugPreviouslyOwnedOwners.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPreviouslyOwnedOwners.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugPreviouslyOwnedOwners.Location = New System.Drawing.Point(0, 0)
        Me.ugPreviouslyOwnedOwners.Name = "ugPreviouslyOwnedOwners"
        Me.ugPreviouslyOwnedOwners.Size = New System.Drawing.Size(954, 299)
        Me.ugPreviouslyOwnedOwners.TabIndex = 98
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.lblNoofOwnersValue)
        Me.Panel3.Controls.Add(Me.lblNoOfOwners)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 299)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(954, 24)
        Me.Panel3.TabIndex = 97
        '
        'lblNoofOwnersValue
        '
        Me.lblNoofOwnersValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoofOwnersValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoofOwnersValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoofOwnersValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNoofOwnersValue.Location = New System.Drawing.Point(123, 0)
        Me.lblNoofOwnersValue.Name = "lblNoofOwnersValue"
        Me.lblNoofOwnersValue.Size = New System.Drawing.Size(36, 24)
        Me.lblNoofOwnersValue.TabIndex = 2
        '
        'lblNoOfOwners
        '
        Me.lblNoOfOwners.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfOwners.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfOwners.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoOfOwners.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNoOfOwners.Location = New System.Drawing.Point(0, 0)
        Me.lblNoOfOwners.Name = "lblNoOfOwners"
        Me.lblNoOfOwners.Size = New System.Drawing.Size(123, 24)
        Me.lblNoOfOwners.TabIndex = 1
        Me.lblNoOfOwners.Text = "No of Previous Owners"
        Me.lblNoOfOwners.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbPageFacilityDocuments
        '
        Me.tbPageFacilityDocuments.Controls.Add(Me.UCFacilityDocuments)
        Me.tbPageFacilityDocuments.Location = New System.Drawing.Point(4, 27)
        Me.tbPageFacilityDocuments.Name = "tbPageFacilityDocuments"
        Me.tbPageFacilityDocuments.Size = New System.Drawing.Size(954, 323)
        Me.tbPageFacilityDocuments.TabIndex = 5
        Me.tbPageFacilityDocuments.Text = "Documents"
        '
        'UCFacilityDocuments
        '
        Me.UCFacilityDocuments.AutoScroll = True
        Me.UCFacilityDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCFacilityDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCFacilityDocuments.Name = "UCFacilityDocuments"
        Me.UCFacilityDocuments.Size = New System.Drawing.Size(954, 323)
        Me.UCFacilityDocuments.TabIndex = 1
        '
        'pnl_FacilityDetail
        '
        Me.pnl_FacilityDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnl_FacilityDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickAssess)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblAssessDate)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkProhibition)
        Me.pnl_FacilityDetail.Controls.Add(Me.LinkLblCAPSignupFac)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacilityLabels)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacilityEnvelopes)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityCompany)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityCompany)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLicensee)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLicenseeSearch)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLicensee)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtDueByNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.Label6)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblActiveLust)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickUpcomingInstallDateValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblUpcomingInstallDate)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkUpcomingInstall)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblNextFac)
        Me.pnl_FacilityDetail.Controls.Add(Me.Panel2)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPStatusValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPStatus)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFuelBrand)
        Me.pnl_FacilityDetail.Controls.Add(Me.ll)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateTransfered)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtFacilityPowerOff)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblPrevFacility)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblPowerOff)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkCAPCandidate)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLocationType)
        Me.pnl_FacilityDetail.Controls.Add(Me.cmbFacilityLocationType)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityMethod)
        Me.pnl_FacilityDetail.Controls.Add(Me.cmbFacilityMethod)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityDatum)
        Me.pnl_FacilityDetail.Controls.Add(Me.cmbFacilityDatum)
        Me.pnl_FacilityDetail.Controls.Add(Me.cmbFacilityType)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLatSec)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLongSec)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLatMin)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLongMin)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLongMin)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLongSec)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatMin)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatSec)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLongDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.mskTxtFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.mskTxtFacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilitySIC)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickFacilityRecvd)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateReceived)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkSignatureofNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilitySigOnNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityFuelBrand)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityStatusValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityStatus)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLongDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLatDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLongitude)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatitude)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityType)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityAIID)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblfacilityAIID)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityIDValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityID)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtfacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityZip)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilitySIC)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityNameForEnsite)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityNameForEnsite)
        Me.pnl_FacilityDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_FacilityDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnl_FacilityDetail.Name = "pnl_FacilityDetail"
        Me.pnl_FacilityDetail.Size = New System.Drawing.Size(964, 280)
        Me.pnl_FacilityDetail.TabIndex = 1
        '
        'dtPickAssess
        '
        Me.dtPickAssess.Checked = False
        Me.dtPickAssess.Enabled = False
        Me.dtPickAssess.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickAssess.Location = New System.Drawing.Point(504, 219)
        Me.dtPickAssess.Name = "dtPickAssess"
        Me.dtPickAssess.ShowCheckBox = True
        Me.dtPickAssess.Size = New System.Drawing.Size(101, 21)
        Me.dtPickAssess.TabIndex = 1059
        '
        'lblAssessDate
        '
        Me.lblAssessDate.Location = New System.Drawing.Point(344, 219)
        Me.lblAssessDate.Name = "lblAssessDate"
        Me.lblAssessDate.Size = New System.Drawing.Size(160, 20)
        Me.lblAssessDate.TabIndex = 1060
        Me.lblAssessDate.Text = "TOS Assessment Date:"
        '
        'chkProhibition
        '
        Me.chkProhibition.CheckAlign = System.Drawing.ContentAlignment.TopRight
        Me.chkProhibition.Location = New System.Drawing.Point(178, 216)
        Me.chkProhibition.Name = "chkProhibition"
        Me.chkProhibition.Size = New System.Drawing.Size(136, 16)
        Me.chkProhibition.TabIndex = 1058
        Me.chkProhibition.Text = "Delivery Prohibition:"
        Me.chkProhibition.Visible = False
        '
        'LinkLblCAPSignupFac
        '
        Me.LinkLblCAPSignupFac.Location = New System.Drawing.Point(800, 240)
        Me.LinkLblCAPSignupFac.Name = "LinkLblCAPSignupFac"
        Me.LinkLblCAPSignupFac.Size = New System.Drawing.Size(152, 16)
        Me.LinkLblCAPSignupFac.TabIndex = 1057
        Me.LinkLblCAPSignupFac.TabStop = True
        Me.LinkLblCAPSignupFac.Text = "CAP Signup/Maintenance"
        '
        'btnFacilityLabels
        '
        Me.btnFacilityLabels.Location = New System.Drawing.Point(8, 119)
        Me.btnFacilityLabels.Name = "btnFacilityLabels"
        Me.btnFacilityLabels.TabIndex = 1056
        Me.btnFacilityLabels.Text = "Labels"
        '
        'btnFacilityEnvelopes
        '
        Me.btnFacilityEnvelopes.Location = New System.Drawing.Point(8, 87)
        Me.btnFacilityEnvelopes.Name = "btnFacilityEnvelopes"
        Me.btnFacilityEnvelopes.TabIndex = 1055
        Me.btnFacilityEnvelopes.Text = "Envelopes"
        '
        'lblFacilityCompany
        '
        Me.lblFacilityCompany.Location = New System.Drawing.Point(616, 208)
        Me.lblFacilityCompany.Name = "lblFacilityCompany"
        Me.lblFacilityCompany.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilityCompany.TabIndex = 1052
        Me.lblFacilityCompany.Text = "Company:"
        '
        'txtFacilityCompany
        '
        Me.txtFacilityCompany.Location = New System.Drawing.Point(704, 208)
        Me.txtFacilityCompany.Name = "txtFacilityCompany"
        Me.txtFacilityCompany.Size = New System.Drawing.Size(144, 21)
        Me.txtFacilityCompany.TabIndex = 1051
        Me.txtFacilityCompany.Text = ""
        '
        'lblFacilityLicensee
        '
        Me.lblFacilityLicensee.Location = New System.Drawing.Point(616, 184)
        Me.lblFacilityLicensee.Name = "lblFacilityLicensee"
        Me.lblFacilityLicensee.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilityLicensee.TabIndex = 1050
        Me.lblFacilityLicensee.Text = "Licensee: "
        '
        'lblFacilityLicenseeSearch
        '
        Me.lblFacilityLicenseeSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFacilityLicenseeSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLicenseeSearch.ForeColor = System.Drawing.Color.Turquoise
        Me.lblFacilityLicenseeSearch.Location = New System.Drawing.Point(856, 184)
        Me.lblFacilityLicenseeSearch.Name = "lblFacilityLicenseeSearch"
        Me.lblFacilityLicenseeSearch.Size = New System.Drawing.Size(16, 23)
        Me.lblFacilityLicenseeSearch.TabIndex = 1049
        Me.lblFacilityLicenseeSearch.Text = "?"
        Me.lblFacilityLicenseeSearch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtFacilityLicensee
        '
        Me.txtFacilityLicensee.Location = New System.Drawing.Point(704, 184)
        Me.txtFacilityLicensee.Name = "txtFacilityLicensee"
        Me.txtFacilityLicensee.Size = New System.Drawing.Size(144, 21)
        Me.txtFacilityLicensee.TabIndex = 1048
        Me.txtFacilityLicensee.Text = ""
        '
        'txtDueByNF
        '
        Me.txtDueByNF.AcceptsTab = True
        Me.txtDueByNF.AutoSize = False
        Me.txtDueByNF.Location = New System.Drawing.Point(928, 160)
        Me.txtDueByNF.Name = "txtDueByNF"
        Me.txtDueByNF.Size = New System.Drawing.Size(104, 21)
        Me.txtDueByNF.TabIndex = 1047
        Me.txtDueByNF.Text = ""
        Me.txtDueByNF.Visible = False
        Me.txtDueByNF.WordWrap = False
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(344, 104)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 20)
        Me.Label6.TabIndex = 1046
        Me.Label6.Text = "Cap Candidate: "
        '
        'lblActiveLust
        '
        Me.lblActiveLust.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblActiveLust.Location = New System.Drawing.Point(344, 56)
        Me.lblActiveLust.Name = "lblActiveLust"
        Me.lblActiveLust.Size = New System.Drawing.Size(104, 20)
        Me.lblActiveLust.TabIndex = 1045
        Me.lblActiveLust.Text = "Active LUST Site: "
        '
        'dtPickUpcomingInstallDateValue
        '
        Me.dtPickUpcomingInstallDateValue.Checked = False
        Me.dtPickUpcomingInstallDateValue.Enabled = False
        Me.dtPickUpcomingInstallDateValue.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickUpcomingInstallDateValue.Location = New System.Drawing.Point(504, 197)
        Me.dtPickUpcomingInstallDateValue.Name = "dtPickUpcomingInstallDateValue"
        Me.dtPickUpcomingInstallDateValue.ShowCheckBox = True
        Me.dtPickUpcomingInstallDateValue.Size = New System.Drawing.Size(101, 21)
        Me.dtPickUpcomingInstallDateValue.TabIndex = 12
        '
        'lblUpcomingInstallDate
        '
        Me.lblUpcomingInstallDate.Location = New System.Drawing.Point(344, 200)
        Me.lblUpcomingInstallDate.Name = "lblUpcomingInstallDate"
        Me.lblUpcomingInstallDate.Size = New System.Drawing.Size(160, 20)
        Me.lblUpcomingInstallDate.TabIndex = 1044
        Me.lblUpcomingInstallDate.Text = "Upcoming Installation Date"
        '
        'chkUpcomingInstall
        '
        Me.chkUpcomingInstall.Location = New System.Drawing.Point(344, 176)
        Me.chkUpcomingInstall.Name = "chkUpcomingInstall"
        Me.chkUpcomingInstall.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkUpcomingInstall.Size = New System.Drawing.Size(145, 24)
        Me.chkUpcomingInstall.TabIndex = 11
        Me.chkUpcomingInstall.Text = "Upcoming Installation "
        Me.chkUpcomingInstall.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lnkLblNextFac
        '
        Me.lnkLblNextFac.Location = New System.Drawing.Point(928, 8)
        Me.lnkLblNextFac.Name = "lnkLblNextFac"
        Me.lnkLblNextFac.Size = New System.Drawing.Size(56, 16)
        Me.lnkLblNextFac.TabIndex = 26
        Me.lnkLblNextFac.TabStop = True
        Me.lnkLblNextFac.Text = "Next>>"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnFacilityCancel)
        Me.Panel2.Controls.Add(Me.btnFacComments)
        Me.Panel2.Controls.Add(Me.btnDeleteFacility)
        Me.Panel2.Controls.Add(Me.btnFacilitySave)
        Me.Panel2.Controls.Add(Me.btnFacFlags)
        Me.Panel2.Location = New System.Drawing.Point(336, 246)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(456, 32)
        Me.Panel2.TabIndex = 24
        '
        'btnFacilityCancel
        '
        Me.btnFacilityCancel.Enabled = False
        Me.btnFacilityCancel.Location = New System.Drawing.Point(108, 4)
        Me.btnFacilityCancel.Name = "btnFacilityCancel"
        Me.btnFacilityCancel.Size = New System.Drawing.Size(75, 26)
        Me.btnFacilityCancel.TabIndex = 1
        Me.btnFacilityCancel.Text = "Cancel"
        '
        'btnFacComments
        '
        Me.btnFacComments.Location = New System.Drawing.Point(365, 4)
        Me.btnFacComments.Name = "btnFacComments"
        Me.btnFacComments.Size = New System.Drawing.Size(80, 26)
        Me.btnFacComments.TabIndex = 4
        Me.btnFacComments.Text = "Comments"
        '
        'btnDeleteFacility
        '
        Me.btnDeleteFacility.Enabled = False
        Me.btnDeleteFacility.Location = New System.Drawing.Point(185, 4)
        Me.btnDeleteFacility.Name = "btnDeleteFacility"
        Me.btnDeleteFacility.Size = New System.Drawing.Size(96, 26)
        Me.btnDeleteFacility.TabIndex = 2
        Me.btnDeleteFacility.Text = "Delete Facility"
        Me.btnDeleteFacility.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnFacilitySave
        '
        Me.btnFacilitySave.Enabled = False
        Me.btnFacilitySave.Location = New System.Drawing.Point(8, 4)
        Me.btnFacilitySave.Name = "btnFacilitySave"
        Me.btnFacilitySave.Size = New System.Drawing.Size(96, 26)
        Me.btnFacilitySave.TabIndex = 0
        Me.btnFacilitySave.Text = "Save Facility"
        Me.btnFacilitySave.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnFacFlags
        '
        Me.btnFacFlags.Location = New System.Drawing.Point(285, 4)
        Me.btnFacFlags.Name = "btnFacFlags"
        Me.btnFacFlags.Size = New System.Drawing.Size(75, 26)
        Me.btnFacFlags.TabIndex = 3
        Me.btnFacFlags.Text = "Flags"
        '
        'lblCAPStatusValue
        '
        Me.lblCAPStatusValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblCAPStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCAPStatusValue.Enabled = False
        Me.lblCAPStatusValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCAPStatusValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblCAPStatusValue.Location = New System.Drawing.Point(456, 80)
        Me.lblCAPStatusValue.Name = "lblCAPStatusValue"
        Me.lblCAPStatusValue.Size = New System.Drawing.Size(120, 20)
        Me.lblCAPStatusValue.TabIndex = 1038
        '
        'lblCAPStatus
        '
        Me.lblCAPStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCAPStatus.Location = New System.Drawing.Point(344, 80)
        Me.lblCAPStatus.Name = "lblCAPStatus"
        Me.lblCAPStatus.Size = New System.Drawing.Size(100, 20)
        Me.lblCAPStatus.TabIndex = 1037
        Me.lblCAPStatus.Text = "CAP Status:"
        '
        'txtFuelBrand
        '
        Me.txtFuelBrand.Location = New System.Drawing.Point(456, 152)
        Me.txtFuelBrand.Name = "txtFuelBrand"
        Me.txtFuelBrand.Size = New System.Drawing.Size(72, 21)
        Me.txtFuelBrand.TabIndex = 10
        Me.txtFuelBrand.Text = ""
        '
        'll
        '
        Me.ll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ll.Location = New System.Drawing.Point(1128, 56)
        Me.ll.Name = "ll"
        Me.ll.Size = New System.Drawing.Size(24, 23)
        Me.ll.TabIndex = 1035
        '
        'lblDateTransfered
        '
        Me.lblDateTransfered.Location = New System.Drawing.Point(896, 152)
        Me.lblDateTransfered.Name = "lblDateTransfered"
        Me.lblDateTransfered.Size = New System.Drawing.Size(32, 16)
        Me.lblDateTransfered.TabIndex = 1034
        Me.lblDateTransfered.Visible = False
        '
        'dtFacilityPowerOff
        '
        Me.dtFacilityPowerOff.Checked = False
        Me.dtFacilityPowerOff.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtFacilityPowerOff.Location = New System.Drawing.Point(704, 192)
        Me.dtFacilityPowerOff.Name = "dtFacilityPowerOff"
        Me.dtFacilityPowerOff.ShowCheckBox = True
        Me.dtFacilityPowerOff.Size = New System.Drawing.Size(104, 21)
        Me.dtFacilityPowerOff.TabIndex = 9
        Me.dtFacilityPowerOff.Visible = False
        '
        'lnkLblPrevFacility
        '
        Me.lnkLblPrevFacility.AutoSize = True
        Me.lnkLblPrevFacility.Location = New System.Drawing.Point(856, 8)
        Me.lnkLblPrevFacility.Name = "lnkLblPrevFacility"
        Me.lnkLblPrevFacility.Size = New System.Drawing.Size(70, 17)
        Me.lnkLblPrevFacility.TabIndex = 25
        Me.lnkLblPrevFacility.TabStop = True
        Me.lnkLblPrevFacility.Text = "<< Previous"
        '
        'chkLUSTSite
        '
        Me.chkLUSTSite.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkLUSTSite.Location = New System.Drawing.Point(456, 56)
        Me.chkLUSTSite.Name = "chkLUSTSite"
        Me.chkLUSTSite.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkLUSTSite.Size = New System.Drawing.Size(16, 20)
        Me.chkLUSTSite.TabIndex = 6
        '
        'lblPowerOff
        '
        Me.lblPowerOff.Location = New System.Drawing.Point(616, 192)
        Me.lblPowerOff.Name = "lblPowerOff"
        Me.lblPowerOff.Size = New System.Drawing.Size(80, 20)
        Me.lblPowerOff.TabIndex = 1028
        Me.lblPowerOff.Text = "Power Off:"
        Me.lblPowerOff.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblPowerOff.Visible = False
        '
        'chkCAPCandidate
        '
        Me.chkCAPCandidate.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkCAPCandidate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCAPCandidate.Location = New System.Drawing.Point(456, 104)
        Me.chkCAPCandidate.Name = "chkCAPCandidate"
        Me.chkCAPCandidate.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkCAPCandidate.Size = New System.Drawing.Size(16, 20)
        Me.chkCAPCandidate.TabIndex = 7
        Me.chkCAPCandidate.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblFacilityLocationType
        '
        Me.lblFacilityLocationType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLocationType.Location = New System.Drawing.Point(616, 152)
        Me.lblFacilityLocationType.Name = "lblFacilityLocationType"
        Me.lblFacilityLocationType.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilityLocationType.TabIndex = 1024
        Me.lblFacilityLocationType.Text = "Type:"
        '
        'cmbFacilityLocationType
        '
        Me.cmbFacilityLocationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityLocationType.DropDownWidth = 250
        Me.cmbFacilityLocationType.Enabled = False
        Me.cmbFacilityLocationType.ItemHeight = 15
        Me.cmbFacilityLocationType.Location = New System.Drawing.Point(704, 152)
        Me.cmbFacilityLocationType.Name = "cmbFacilityLocationType"
        Me.cmbFacilityLocationType.Size = New System.Drawing.Size(144, 23)
        Me.cmbFacilityLocationType.TabIndex = 23
        '
        'lblFacilityMethod
        '
        Me.lblFacilityMethod.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityMethod.Location = New System.Drawing.Point(616, 128)
        Me.lblFacilityMethod.Name = "lblFacilityMethod"
        Me.lblFacilityMethod.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityMethod.TabIndex = 1022
        Me.lblFacilityMethod.Text = "Method:"
        '
        'cmbFacilityMethod
        '
        Me.cmbFacilityMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityMethod.DropDownWidth = 350
        Me.cmbFacilityMethod.Enabled = False
        Me.cmbFacilityMethod.ItemHeight = 15
        Me.cmbFacilityMethod.Location = New System.Drawing.Point(704, 128)
        Me.cmbFacilityMethod.Name = "cmbFacilityMethod"
        Me.cmbFacilityMethod.Size = New System.Drawing.Size(144, 23)
        Me.cmbFacilityMethod.TabIndex = 22
        '
        'lblFacilityDatum
        '
        Me.lblFacilityDatum.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityDatum.Location = New System.Drawing.Point(616, 104)
        Me.lblFacilityDatum.Name = "lblFacilityDatum"
        Me.lblFacilityDatum.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityDatum.TabIndex = 1020
        Me.lblFacilityDatum.Text = "Datum:"
        '
        'cmbFacilityDatum
        '
        Me.cmbFacilityDatum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityDatum.DropDownWidth = 250
        Me.cmbFacilityDatum.Enabled = False
        Me.cmbFacilityDatum.ItemHeight = 15
        Me.cmbFacilityDatum.Location = New System.Drawing.Point(704, 104)
        Me.cmbFacilityDatum.Name = "cmbFacilityDatum"
        Me.cmbFacilityDatum.Size = New System.Drawing.Size(144, 23)
        Me.cmbFacilityDatum.TabIndex = 21
        '
        'cmbFacilityType
        '
        Me.cmbFacilityType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityType.DropDownWidth = 180
        Me.cmbFacilityType.ItemHeight = 15
        Me.cmbFacilityType.Location = New System.Drawing.Point(704, 32)
        Me.cmbFacilityType.Name = "cmbFacilityType"
        Me.cmbFacilityType.Size = New System.Drawing.Size(144, 23)
        Me.cmbFacilityType.TabIndex = 14
        '
        'txtFacilityLatSec
        '
        Me.txtFacilityLatSec.AcceptsTab = True
        Me.txtFacilityLatSec.AutoSize = False
        Me.txtFacilityLatSec.Location = New System.Drawing.Point(800, 56)
        Me.txtFacilityLatSec.MaxLength = 5
        Me.txtFacilityLatSec.Name = "txtFacilityLatSec"
        Me.txtFacilityLatSec.Size = New System.Drawing.Size(37, 20)
        Me.txtFacilityLatSec.TabIndex = 17
        Me.txtFacilityLatSec.Text = ""
        Me.txtFacilityLatSec.WordWrap = False
        '
        'txtFacilityLongSec
        '
        Me.txtFacilityLongSec.AcceptsTab = True
        Me.txtFacilityLongSec.AutoSize = False
        Me.txtFacilityLongSec.Location = New System.Drawing.Point(800, 80)
        Me.txtFacilityLongSec.MaxLength = 5
        Me.txtFacilityLongSec.Name = "txtFacilityLongSec"
        Me.txtFacilityLongSec.Size = New System.Drawing.Size(38, 20)
        Me.txtFacilityLongSec.TabIndex = 20
        Me.txtFacilityLongSec.Text = ""
        Me.txtFacilityLongSec.WordWrap = False
        '
        'txtFacilityLatMin
        '
        Me.txtFacilityLatMin.AcceptsTab = True
        Me.txtFacilityLatMin.AutoSize = False
        Me.txtFacilityLatMin.Location = New System.Drawing.Point(754, 56)
        Me.txtFacilityLatMin.MaxLength = 2
        Me.txtFacilityLatMin.Name = "txtFacilityLatMin"
        Me.txtFacilityLatMin.Size = New System.Drawing.Size(28, 20)
        Me.txtFacilityLatMin.TabIndex = 16
        Me.txtFacilityLatMin.Text = ""
        Me.txtFacilityLatMin.WordWrap = False
        '
        'txtFacilityLongMin
        '
        Me.txtFacilityLongMin.AcceptsTab = True
        Me.txtFacilityLongMin.AutoSize = False
        Me.txtFacilityLongMin.Location = New System.Drawing.Point(754, 80)
        Me.txtFacilityLongMin.MaxLength = 2
        Me.txtFacilityLongMin.Name = "txtFacilityLongMin"
        Me.txtFacilityLongMin.Size = New System.Drawing.Size(28, 20)
        Me.txtFacilityLongMin.TabIndex = 19
        Me.txtFacilityLongMin.Text = ""
        Me.txtFacilityLongMin.WordWrap = False
        '
        'lblFacilityLongMin
        '
        Me.lblFacilityLongMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongMin.Location = New System.Drawing.Point(784, 80)
        Me.lblFacilityLongMin.Name = "lblFacilityLongMin"
        Me.lblFacilityLongMin.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLongMin.TabIndex = 1018
        Me.lblFacilityLongMin.Text = "'"
        Me.lblFacilityLongMin.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblFacilityLongSec
        '
        Me.lblFacilityLongSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongSec.Location = New System.Drawing.Point(840, 80)
        Me.lblFacilityLongSec.Name = "lblFacilityLongSec"
        Me.lblFacilityLongSec.Size = New System.Drawing.Size(10, 16)
        Me.lblFacilityLongSec.TabIndex = 1017
        Me.lblFacilityLongSec.Text = """"
        '
        'lblFacilityLatMin
        '
        Me.lblFacilityLatMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatMin.Location = New System.Drawing.Point(784, 56)
        Me.lblFacilityLatMin.Name = "lblFacilityLatMin"
        Me.lblFacilityLatMin.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLatMin.TabIndex = 1016
        Me.lblFacilityLatMin.Text = "'"
        Me.lblFacilityLatMin.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblFacilityLatSec
        '
        Me.lblFacilityLatSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatSec.Location = New System.Drawing.Point(840, 56)
        Me.lblFacilityLatSec.Name = "lblFacilityLatSec"
        Me.lblFacilityLatSec.Size = New System.Drawing.Size(10, 16)
        Me.lblFacilityLatSec.TabIndex = 1015
        Me.lblFacilityLatSec.Text = """"
        '
        'lblFacilityLongDegree
        '
        Me.lblFacilityLongDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongDegree.Location = New System.Drawing.Point(736, 74)
        Me.lblFacilityLongDegree.Name = "lblFacilityLongDegree"
        Me.lblFacilityLongDegree.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLongDegree.TabIndex = 1010
        Me.lblFacilityLongDegree.Text = "o"
        Me.lblFacilityLongDegree.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'mskTxtFacilityFax
        '
        Me.mskTxtFacilityFax.ContainingControl = Me
        Me.mskTxtFacilityFax.Location = New System.Drawing.Point(88, 186)
        Me.mskTxtFacilityFax.Name = "mskTxtFacilityFax"
        Me.mskTxtFacilityFax.OcxState = CType(resources.GetObject("mskTxtFacilityFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFacilityFax.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtFacilityFax.TabIndex = 3
        '
        'mskTxtFacilityPhone
        '
        Me.mskTxtFacilityPhone.ContainingControl = Me
        Me.mskTxtFacilityPhone.Location = New System.Drawing.Point(88, 162)
        Me.mskTxtFacilityPhone.Name = "mskTxtFacilityPhone"
        Me.mskTxtFacilityPhone.OcxState = CType(resources.GetObject("mskTxtFacilityPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFacilityPhone.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtFacilityPhone.TabIndex = 2
        '
        'txtFacilityAddress
        '
        Me.txtFacilityAddress.Location = New System.Drawing.Point(88, 56)
        Me.txtFacilityAddress.Multiline = True
        Me.txtFacilityAddress.Name = "txtFacilityAddress"
        Me.txtFacilityAddress.ReadOnly = True
        Me.txtFacilityAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtFacilityAddress.Size = New System.Drawing.Size(224, 104)
        Me.txtFacilityAddress.TabIndex = 1
        Me.txtFacilityAddress.Text = ""
        Me.txtFacilityAddress.WordWrap = False
        '
        'lblFacilitySIC
        '
        Me.lblFacilitySIC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySIC.Location = New System.Drawing.Point(344, 128)
        Me.lblFacilitySIC.Name = "lblFacilitySIC"
        Me.lblFacilitySIC.Size = New System.Drawing.Size(100, 20)
        Me.lblFacilitySIC.TabIndex = 150
        Me.lblFacilitySIC.Text = "SIC:"
        '
        'txtFacilityFax
        '
        Me.txtFacilityFax.AcceptsTab = True
        Me.txtFacilityFax.AutoSize = False
        Me.txtFacilityFax.Location = New System.Drawing.Point(928, 208)
        Me.txtFacilityFax.Name = "txtFacilityFax"
        Me.txtFacilityFax.Size = New System.Drawing.Size(104, 21)
        Me.txtFacilityFax.TabIndex = 148
        Me.txtFacilityFax.Text = ""
        Me.txtFacilityFax.Visible = False
        Me.txtFacilityFax.WordWrap = False
        '
        'lblFacilityFax
        '
        Me.lblFacilityFax.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityFax.Location = New System.Drawing.Point(8, 186)
        Me.lblFacilityFax.Name = "lblFacilityFax"
        Me.lblFacilityFax.Size = New System.Drawing.Size(72, 23)
        Me.lblFacilityFax.TabIndex = 147
        Me.lblFacilityFax.Text = "Fax:"
        '
        'dtPickFacilityRecvd
        '
        Me.dtPickFacilityRecvd.Checked = False
        Me.dtPickFacilityRecvd.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickFacilityRecvd.Location = New System.Drawing.Point(456, 8)
        Me.dtPickFacilityRecvd.Name = "dtPickFacilityRecvd"
        Me.dtPickFacilityRecvd.ShowCheckBox = True
        Me.dtPickFacilityRecvd.Size = New System.Drawing.Size(104, 21)
        Me.dtPickFacilityRecvd.TabIndex = 5
        '
        'lblDateReceived
        '
        Me.lblDateReceived.Location = New System.Drawing.Point(344, 8)
        Me.lblDateReceived.Name = "lblDateReceived"
        Me.lblDateReceived.Size = New System.Drawing.Size(100, 20)
        Me.lblDateReceived.TabIndex = 145
        Me.lblDateReceived.Text = "Date Received:"
        '
        'chkSignatureofNF
        '
        Me.chkSignatureofNF.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkSignatureofNF.Location = New System.Drawing.Point(6, 216)
        Me.chkSignatureofNF.Name = "chkSignatureofNF"
        Me.chkSignatureofNF.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkSignatureofNF.Size = New System.Drawing.Size(144, 20)
        Me.chkSignatureofNF.TabIndex = 4
        Me.chkSignatureofNF.Text = ": Signature Received"
        Me.chkSignatureofNF.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFacilitySigOnNF
        '
        Me.lblFacilitySigOnNF.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySigOnNF.Location = New System.Drawing.Point(8, 216)
        Me.lblFacilitySigOnNF.Name = "lblFacilitySigOnNF"
        Me.lblFacilitySigOnNF.Size = New System.Drawing.Size(120, 23)
        Me.lblFacilitySigOnNF.TabIndex = 127
        Me.lblFacilitySigOnNF.Text = "Signature Received:"
        '
        'lblFacilityFuelBrand
        '
        Me.lblFacilityFuelBrand.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityFuelBrand.Location = New System.Drawing.Point(344, 152)
        Me.lblFacilityFuelBrand.Name = "lblFacilityFuelBrand"
        Me.lblFacilityFuelBrand.Size = New System.Drawing.Size(100, 20)
        Me.lblFacilityFuelBrand.TabIndex = 125
        Me.lblFacilityFuelBrand.Text = "Fuel Brand:"
        '
        'lblFacilityStatusValue
        '
        Me.lblFacilityStatusValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblFacilityStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFacilityStatusValue.Enabled = False
        Me.lblFacilityStatusValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityStatusValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblFacilityStatusValue.Location = New System.Drawing.Point(456, 32)
        Me.lblFacilityStatusValue.Name = "lblFacilityStatusValue"
        Me.lblFacilityStatusValue.Size = New System.Drawing.Size(120, 20)
        Me.lblFacilityStatusValue.TabIndex = 124
        '
        'lblFacilityStatus
        '
        Me.lblFacilityStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityStatus.Location = New System.Drawing.Point(344, 32)
        Me.lblFacilityStatus.Name = "lblFacilityStatus"
        Me.lblFacilityStatus.Size = New System.Drawing.Size(100, 20)
        Me.lblFacilityStatus.TabIndex = 123
        Me.lblFacilityStatus.Text = "Facility Status:"
        '
        'txtFacilityLongDegree
        '
        Me.txtFacilityLongDegree.AcceptsTab = True
        Me.txtFacilityLongDegree.AutoSize = False
        Me.txtFacilityLongDegree.Location = New System.Drawing.Point(704, 80)
        Me.txtFacilityLongDegree.MaxLength = 3
        Me.txtFacilityLongDegree.Name = "txtFacilityLongDegree"
        Me.txtFacilityLongDegree.Size = New System.Drawing.Size(32, 20)
        Me.txtFacilityLongDegree.TabIndex = 18
        Me.txtFacilityLongDegree.Text = ""
        Me.txtFacilityLongDegree.WordWrap = False
        '
        'txtFacilityLatDegree
        '
        Me.txtFacilityLatDegree.AcceptsTab = True
        Me.txtFacilityLatDegree.AutoSize = False
        Me.txtFacilityLatDegree.Location = New System.Drawing.Point(704, 56)
        Me.txtFacilityLatDegree.MaxLength = 3
        Me.txtFacilityLatDegree.Name = "txtFacilityLatDegree"
        Me.txtFacilityLatDegree.Size = New System.Drawing.Size(32, 20)
        Me.txtFacilityLatDegree.TabIndex = 15
        Me.txtFacilityLatDegree.Text = ""
        Me.txtFacilityLatDegree.WordWrap = False
        '
        'lblFacilityLongitude
        '
        Me.lblFacilityLongitude.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongitude.Location = New System.Drawing.Point(616, 80)
        Me.lblFacilityLongitude.Name = "lblFacilityLongitude"
        Me.lblFacilityLongitude.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityLongitude.TabIndex = 111
        Me.lblFacilityLongitude.Text = "Longitude:"
        '
        'lblFacilityLatitude
        '
        Me.lblFacilityLatitude.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatitude.Location = New System.Drawing.Point(616, 56)
        Me.lblFacilityLatitude.Name = "lblFacilityLatitude"
        Me.lblFacilityLatitude.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityLatitude.TabIndex = 110
        Me.lblFacilityLatitude.Text = "Latitude:"
        '
        'lblFacilityType
        '
        Me.lblFacilityType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityType.Location = New System.Drawing.Point(616, 32)
        Me.lblFacilityType.Name = "lblFacilityType"
        Me.lblFacilityType.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityType.TabIndex = 106
        Me.lblFacilityType.Text = "Facility Type:"
        '
        'txtFacilityAIID
        '
        Me.txtFacilityAIID.AcceptsTab = True
        Me.txtFacilityAIID.AutoSize = False
        Me.txtFacilityAIID.Location = New System.Drawing.Point(704, 8)
        Me.txtFacilityAIID.Name = "txtFacilityAIID"
        Me.txtFacilityAIID.ReadOnly = True
        Me.txtFacilityAIID.Size = New System.Drawing.Size(144, 20)
        Me.txtFacilityAIID.TabIndex = 13
        Me.txtFacilityAIID.Text = ""
        Me.txtFacilityAIID.WordWrap = False
        '
        'lblfacilityAIID
        '
        Me.lblfacilityAIID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblfacilityAIID.Location = New System.Drawing.Point(616, 8)
        Me.lblfacilityAIID.Name = "lblfacilityAIID"
        Me.lblfacilityAIID.Size = New System.Drawing.Size(80, 20)
        Me.lblfacilityAIID.TabIndex = 104
        Me.lblfacilityAIID.Text = "Facility AIID:"
        '
        'lblFacilityIDValue
        '
        Me.lblFacilityIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityIDValue.Location = New System.Drawing.Point(88, 8)
        Me.lblFacilityIDValue.Name = "lblFacilityIDValue"
        Me.lblFacilityIDValue.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityIDValue.TabIndex = 103
        '
        'lblFacilityID
        '
        Me.lblFacilityID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityID.Location = New System.Drawing.Point(8, 8)
        Me.lblFacilityID.Name = "lblFacilityID"
        Me.lblFacilityID.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityID.TabIndex = 102
        Me.lblFacilityID.Text = "Facility ID:"
        '
        'txtfacilityPhone
        '
        Me.txtfacilityPhone.AcceptsTab = True
        Me.txtfacilityPhone.AutoSize = False
        Me.txtfacilityPhone.Location = New System.Drawing.Point(928, 184)
        Me.txtfacilityPhone.Name = "txtfacilityPhone"
        Me.txtfacilityPhone.Size = New System.Drawing.Size(104, 21)
        Me.txtfacilityPhone.TabIndex = 99
        Me.txtfacilityPhone.Text = ""
        Me.txtfacilityPhone.Visible = False
        Me.txtfacilityPhone.WordWrap = False
        '
        'lblFacilityPhone
        '
        Me.lblFacilityPhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityPhone.Location = New System.Drawing.Point(8, 162)
        Me.lblFacilityPhone.Name = "lblFacilityPhone"
        Me.lblFacilityPhone.Size = New System.Drawing.Size(72, 23)
        Me.lblFacilityPhone.TabIndex = 98
        Me.lblFacilityPhone.Text = "Phone:"
        '
        'txtFacilityName
        '
        Me.txtFacilityName.AcceptsTab = True
        Me.txtFacilityName.AutoSize = False
        Me.txtFacilityName.Location = New System.Drawing.Point(88, 32)
        Me.txtFacilityName.Name = "txtFacilityName"
        Me.txtFacilityName.Size = New System.Drawing.Size(224, 21)
        Me.txtFacilityName.TabIndex = 0
        Me.txtFacilityName.Text = ""
        Me.txtFacilityName.WordWrap = False
        '
        'lblFacilityAddress
        '
        Me.lblFacilityAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityAddress.Location = New System.Drawing.Point(8, 56)
        Me.lblFacilityAddress.Name = "lblFacilityAddress"
        Me.lblFacilityAddress.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityAddress.TabIndex = 90
        Me.lblFacilityAddress.Text = "Address:"
        '
        'lblFacilityName
        '
        Me.lblFacilityName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityName.Location = New System.Drawing.Point(8, 32)
        Me.lblFacilityName.Name = "lblFacilityName"
        Me.lblFacilityName.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityName.TabIndex = 89
        Me.lblFacilityName.Text = "Facility Name:"
        '
        'txtFacilityZip
        '
        Me.txtFacilityZip.Location = New System.Drawing.Point(896, 32)
        Me.txtFacilityZip.Name = "txtFacilityZip"
        Me.txtFacilityZip.Size = New System.Drawing.Size(96, 21)
        Me.txtFacilityZip.TabIndex = 11
        Me.txtFacilityZip.Text = ""
        Me.txtFacilityZip.Visible = False
        '
        'lblFacilityLatDegree
        '
        Me.lblFacilityLatDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatDegree.Location = New System.Drawing.Point(736, 54)
        Me.lblFacilityLatDegree.Name = "lblFacilityLatDegree"
        Me.lblFacilityLatDegree.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLatDegree.TabIndex = 1009
        Me.lblFacilityLatDegree.Text = "o"
        Me.lblFacilityLatDegree.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtFacilitySIC
        '
        Me.txtFacilitySIC.BackColor = System.Drawing.SystemColors.Control
        Me.txtFacilitySIC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFacilitySIC.Enabled = False
        Me.txtFacilitySIC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFacilitySIC.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.txtFacilitySIC.Location = New System.Drawing.Point(456, 128)
        Me.txtFacilitySIC.Name = "txtFacilitySIC"
        Me.txtFacilitySIC.Size = New System.Drawing.Size(120, 20)
        Me.txtFacilitySIC.TabIndex = 1038
        '
        'txtFacilityNameForEnsite
        '
        Me.txtFacilityNameForEnsite.AcceptsTab = True
        Me.txtFacilityNameForEnsite.AutoSize = False
        Me.txtFacilityNameForEnsite.Location = New System.Drawing.Point(88, 240)
        Me.txtFacilityNameForEnsite.Name = "txtFacilityNameForEnsite"
        Me.txtFacilityNameForEnsite.Size = New System.Drawing.Size(224, 21)
        Me.txtFacilityNameForEnsite.TabIndex = 5
        Me.txtFacilityNameForEnsite.Text = ""
        Me.txtFacilityNameForEnsite.WordWrap = False
        '
        'lblFacilityNameForEnsite
        '
        Me.lblFacilityNameForEnsite.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityNameForEnsite.Location = New System.Drawing.Point(8, 240)
        Me.lblFacilityNameForEnsite.Name = "lblFacilityNameForEnsite"
        Me.lblFacilityNameForEnsite.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityNameForEnsite.TabIndex = 89
        Me.lblFacilityNameForEnsite.Text = "Ensite Name:"
        '
        'tbPageManageTank
        '
        Me.tbPageManageTank.AutoScroll = True
        Me.tbPageManageTank.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageManageTank.Controls.Add(Me.dgPipesAndTanks2)
        Me.tbPageManageTank.Controls.Add(Me.tbCntrlPipe)
        Me.tbPageManageTank.Controls.Add(Me.tbCntrlTank)
        Me.tbPageManageTank.Controls.Add(Me.pnlTankDetailMainDisplay)
        Me.tbPageManageTank.Controls.Add(Me.pnlTankCount2)
        Me.tbPageManageTank.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPageManageTank.Location = New System.Drawing.Point(4, 22)
        Me.tbPageManageTank.Name = "tbPageManageTank"
        Me.tbPageManageTank.Size = New System.Drawing.Size(964, 636)
        Me.tbPageManageTank.TabIndex = 9
        Me.tbPageManageTank.Text = "Tank/Pipe Summary"
        '
        'dgPipesAndTanks2
        '
        Me.dgPipesAndTanks2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.dgPipesAndTanks2.Dock = System.Windows.Forms.DockStyle.Top
        Me.dgPipesAndTanks2.Location = New System.Drawing.Point(0, 32)
        Me.dgPipesAndTanks2.Name = "dgPipesAndTanks2"
        Me.dgPipesAndTanks2.Size = New System.Drawing.Size(960, 160)
        Me.dgPipesAndTanks2.TabIndex = 11
        Me.dgPipesAndTanks2.Text = "Tanks And Pipes"
        '
        'tbCntrlPipe
        '
        Me.tbCntrlPipe.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.tbCntrlPipe.Controls.Add(Me.tbPagePipeDetail)
        Me.tbCntrlPipe.Location = New System.Drawing.Point(0, 208)
        Me.tbCntrlPipe.Name = "tbCntrlPipe"
        Me.tbCntrlPipe.SelectedIndex = 0
        Me.tbCntrlPipe.Size = New System.Drawing.Size(864, 208)
        Me.tbCntrlPipe.TabIndex = 1
        Me.tbCntrlPipe.Visible = False
        '
        'tbPagePipeDetail
        '
        Me.tbPagePipeDetail.Controls.Add(Me.pnlPipeDetail)
        Me.tbPagePipeDetail.Controls.Add(Me.pnlPipeButtons)
        Me.tbPagePipeDetail.Controls.Add(Me.pnlPipeDetailHeader)
        Me.tbPagePipeDetail.Location = New System.Drawing.Point(4, 27)
        Me.tbPagePipeDetail.Name = "tbPagePipeDetail"
        Me.tbPagePipeDetail.Size = New System.Drawing.Size(856, 177)
        Me.tbPagePipeDetail.TabIndex = 0
        Me.tbPagePipeDetail.Text = "Pipe Detail"
        '
        'pnlPipeDetail
        '
        Me.pnlPipeDetail.AutoScroll = True
        Me.pnlPipeDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeClosure)
        Me.pnlPipeDetail.Controls.Add(Me.Panel20)
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeInstallerOath)
        Me.pnlPipeDetail.Controls.Add(Me.Panel9)
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeRelease)
        Me.pnlPipeDetail.Controls.Add(Me.Panel11)
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeType)
        Me.pnlPipeDetail.Controls.Add(Me.Panel15)
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeMaterial)
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeMaterialHead)
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeDateOfInstallation)
        Me.pnlPipeDetail.Controls.Add(Me.Panel13)
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeDescription)
        Me.pnlPipeDetail.Controls.Add(Me.pnlPipeDescHead)
        Me.pnlPipeDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlPipeDetail.DockPadding.All = 3
        Me.pnlPipeDetail.Location = New System.Drawing.Point(0, 24)
        Me.pnlPipeDetail.Name = "pnlPipeDetail"
        Me.pnlPipeDetail.Size = New System.Drawing.Size(856, 113)
        Me.pnlPipeDetail.TabIndex = 0
        '
        'pnlPipeClosure
        '
        Me.pnlPipeClosure.Controls.Add(Me.cmbPipeInertFill)
        Me.pnlPipeClosure.Controls.Add(Me.cmbPipeClosureType)
        Me.pnlPipeClosure.Controls.Add(Me.dtPickPipeLastUsed)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeLastUsed)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeClosedOn)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeClosedOnDate)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeInertFillValue)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeInertFill)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeClosureStatusValue)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeClosureStatus)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeClosureRcvdDateValue)
        Me.pnlPipeClosure.Controls.Add(Me.lblPipeClosureRcvd)
        Me.pnlPipeClosure.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeClosure.Location = New System.Drawing.Point(3, 968)
        Me.pnlPipeClosure.Name = "pnlPipeClosure"
        Me.pnlPipeClosure.Size = New System.Drawing.Size(846, 80)
        Me.pnlPipeClosure.TabIndex = 6
        '
        'cmbPipeInertFill
        '
        Me.cmbPipeInertFill.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeInertFill.DropDownWidth = 300
        Me.cmbPipeInertFill.Location = New System.Drawing.Point(462, 32)
        Me.cmbPipeInertFill.Name = "cmbPipeInertFill"
        Me.cmbPipeInertFill.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeInertFill.TabIndex = 2
        '
        'cmbPipeClosureType
        '
        Me.cmbPipeClosureType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeClosureType.DropDownWidth = 300
        Me.cmbPipeClosureType.Location = New System.Drawing.Point(462, 8)
        Me.cmbPipeClosureType.Name = "cmbPipeClosureType"
        Me.cmbPipeClosureType.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeClosureType.TabIndex = 1
        '
        'dtPickPipeLastUsed
        '
        Me.dtPickPipeLastUsed.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickPipeLastUsed.Checked = False
        Me.dtPickPipeLastUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickPipeLastUsed.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeLastUsed.Location = New System.Drawing.Point(136, 8)
        Me.dtPickPipeLastUsed.Name = "dtPickPipeLastUsed"
        Me.dtPickPipeLastUsed.ShowCheckBox = True
        Me.dtPickPipeLastUsed.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeLastUsed.TabIndex = 0
        Me.dtPickPipeLastUsed.Value = New Date(2004, 5, 27, 14, 25, 14, 144)
        '
        'lblPipeLastUsed
        '
        Me.lblPipeLastUsed.AutoSize = True
        Me.lblPipeLastUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeLastUsed.Location = New System.Drawing.Point(40, 8)
        Me.lblPipeLastUsed.Name = "lblPipeLastUsed"
        Me.lblPipeLastUsed.Size = New System.Drawing.Size(93, 17)
        Me.lblPipeLastUsed.TabIndex = 183
        Me.lblPipeLastUsed.Text = "Date Last Used:"
        '
        'lblPipeClosedOn
        '
        Me.lblPipeClosedOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeClosedOn.Location = New System.Drawing.Point(54, 54)
        Me.lblPipeClosedOn.Name = "lblPipeClosedOn"
        Me.lblPipeClosedOn.Size = New System.Drawing.Size(80, 23)
        Me.lblPipeClosedOn.TabIndex = 184
        Me.lblPipeClosedOn.Text = "Date Closed:"
        '
        'lblPipeClosedOnDate
        '
        Me.lblPipeClosedOnDate.AutoSize = True
        Me.lblPipeClosedOnDate.BackColor = System.Drawing.Color.Red
        Me.lblPipeClosedOnDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeClosedOnDate.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblPipeClosedOnDate.Location = New System.Drawing.Point(135, 55)
        Me.lblPipeClosedOnDate.Name = "lblPipeClosedOnDate"
        Me.lblPipeClosedOnDate.Size = New System.Drawing.Size(129, 17)
        Me.lblPipeClosedOnDate.TabIndex = 185
        Me.lblPipeClosedOnDate.Text = "If Closed,Closure Date"
        '
        'lblPipeInertFillValue
        '
        Me.lblPipeInertFillValue.Location = New System.Drawing.Point(576, 33)
        Me.lblPipeInertFillValue.Name = "lblPipeInertFillValue"
        Me.lblPipeInertFillValue.TabIndex = 182
        Me.lblPipeInertFillValue.Text = "Concrete"
        '
        'lblPipeInertFill
        '
        Me.lblPipeInertFill.Location = New System.Drawing.Point(407, 33)
        Me.lblPipeInertFill.Name = "lblPipeInertFill"
        Me.lblPipeInertFill.Size = New System.Drawing.Size(49, 23)
        Me.lblPipeInertFill.TabIndex = 181
        Me.lblPipeInertFill.Text = "InertFill:"
        '
        'lblPipeClosureStatusValue
        '
        Me.lblPipeClosureStatusValue.Location = New System.Drawing.Point(576, 9)
        Me.lblPipeClosureStatusValue.Name = "lblPipeClosureStatusValue"
        Me.lblPipeClosureStatusValue.TabIndex = 180
        Me.lblPipeClosureStatusValue.Text = "Closed"
        '
        'lblPipeClosureStatus
        '
        Me.lblPipeClosureStatus.Location = New System.Drawing.Point(375, 9)
        Me.lblPipeClosureStatus.Name = "lblPipeClosureStatus"
        Me.lblPipeClosureStatus.Size = New System.Drawing.Size(81, 21)
        Me.lblPipeClosureStatus.TabIndex = 179
        Me.lblPipeClosureStatus.Text = "Closure Type:"
        '
        'lblPipeClosureRcvdDateValue
        '
        Me.lblPipeClosureRcvdDateValue.Location = New System.Drawing.Point(136, 32)
        Me.lblPipeClosureRcvdDateValue.Name = "lblPipeClosureRcvdDateValue"
        Me.lblPipeClosureRcvdDateValue.TabIndex = 178
        '
        'lblPipeClosureRcvd
        '
        Me.lblPipeClosureRcvd.Location = New System.Drawing.Point(20, 32)
        Me.lblPipeClosureRcvd.Name = "lblPipeClosureRcvd"
        Me.lblPipeClosureRcvd.Size = New System.Drawing.Size(120, 16)
        Me.lblPipeClosureRcvd.TabIndex = 177
        Me.lblPipeClosureRcvd.Text = "Date Closure Rcvd:"
        '
        'Panel20
        '
        Me.Panel20.Controls.Add(Me.lblPipeClosureCaption)
        Me.Panel20.Controls.Add(Me.lblPipeClosure)
        Me.Panel20.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel20.Location = New System.Drawing.Point(3, 944)
        Me.Panel20.Name = "Panel20"
        Me.Panel20.Size = New System.Drawing.Size(846, 24)
        Me.Panel20.TabIndex = 193
        '
        'lblPipeClosureCaption
        '
        Me.lblPipeClosureCaption.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPipeClosureCaption.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPipeClosureCaption.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPipeClosureCaption.Location = New System.Drawing.Point(16, 0)
        Me.lblPipeClosureCaption.Name = "lblPipeClosureCaption"
        Me.lblPipeClosureCaption.Size = New System.Drawing.Size(830, 24)
        Me.lblPipeClosureCaption.TabIndex = 3
        Me.lblPipeClosureCaption.Text = "Closing of Pipe"
        Me.lblPipeClosureCaption.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPipeClosure
        '
        Me.lblPipeClosure.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPipeClosure.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPipeClosure.Location = New System.Drawing.Point(0, 0)
        Me.lblPipeClosure.Name = "lblPipeClosure"
        Me.lblPipeClosure.Size = New System.Drawing.Size(16, 24)
        Me.lblPipeClosure.TabIndex = 2
        Me.lblPipeClosure.Text = "-"
        Me.lblPipeClosure.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPipeInstallerOath
        '
        Me.pnlPipeInstallerOath.Controls.Add(Me.txtPipeLicensee)
        Me.pnlPipeInstallerOath.Controls.Add(Me.lblPipeLicenseeCompanySearch)
        Me.pnlPipeInstallerOath.Controls.Add(Me.txtPipeCompanyName)
        Me.pnlPipeInstallerOath.Controls.Add(Me.lblPipeInstallerName)
        Me.pnlPipeInstallerOath.Controls.Add(Me.lblPipeInstallerCompanyName)
        Me.pnlPipeInstallerOath.Controls.Add(Me.dtPickPipeSigned)
        Me.pnlPipeInstallerOath.Controls.Add(Me.lblDtPipeSigned)
        Me.pnlPipeInstallerOath.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeInstallerOath.Location = New System.Drawing.Point(3, 856)
        Me.pnlPipeInstallerOath.Name = "pnlPipeInstallerOath"
        Me.pnlPipeInstallerOath.Size = New System.Drawing.Size(846, 88)
        Me.pnlPipeInstallerOath.TabIndex = 5
        '
        'txtPipeLicensee
        '
        Me.txtPipeLicensee.Location = New System.Drawing.Point(184, 32)
        Me.txtPipeLicensee.Name = "txtPipeLicensee"
        Me.txtPipeLicensee.Size = New System.Drawing.Size(200, 21)
        Me.txtPipeLicensee.TabIndex = 222
        Me.txtPipeLicensee.Text = ""
        '
        'lblPipeLicenseeCompanySearch
        '
        Me.lblPipeLicenseeCompanySearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPipeLicenseeCompanySearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeLicenseeCompanySearch.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblPipeLicenseeCompanySearch.Location = New System.Drawing.Point(392, 32)
        Me.lblPipeLicenseeCompanySearch.Name = "lblPipeLicenseeCompanySearch"
        Me.lblPipeLicenseeCompanySearch.Size = New System.Drawing.Size(16, 22)
        Me.lblPipeLicenseeCompanySearch.TabIndex = 221
        Me.lblPipeLicenseeCompanySearch.Text = "?"
        Me.lblPipeLicenseeCompanySearch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtPipeCompanyName
        '
        Me.txtPipeCompanyName.Location = New System.Drawing.Point(184, 56)
        Me.txtPipeCompanyName.Name = "txtPipeCompanyName"
        Me.txtPipeCompanyName.Size = New System.Drawing.Size(200, 21)
        Me.txtPipeCompanyName.TabIndex = 195
        Me.txtPipeCompanyName.Text = ""
        '
        'lblPipeInstallerName
        '
        Me.lblPipeInstallerName.Location = New System.Drawing.Point(88, 32)
        Me.lblPipeInstallerName.Name = "lblPipeInstallerName"
        Me.lblPipeInstallerName.Size = New System.Drawing.Size(88, 17)
        Me.lblPipeInstallerName.TabIndex = 192
        Me.lblPipeInstallerName.Text = "Installer Name:"
        Me.lblPipeInstallerName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPipeInstallerCompanyName
        '
        Me.lblPipeInstallerCompanyName.AutoSize = True
        Me.lblPipeInstallerCompanyName.Location = New System.Drawing.Point(112, 56)
        Me.lblPipeInstallerCompanyName.Name = "lblPipeInstallerCompanyName"
        Me.lblPipeInstallerCompanyName.Size = New System.Drawing.Size(61, 17)
        Me.lblPipeInstallerCompanyName.TabIndex = 193
        Me.lblPipeInstallerCompanyName.Text = "Company:"
        '
        'dtPickPipeSigned
        '
        Me.dtPickPipeSigned.Checked = False
        Me.dtPickPipeSigned.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeSigned.Location = New System.Drawing.Point(184, 4)
        Me.dtPickPipeSigned.Name = "dtPickPipeSigned"
        Me.dtPickPipeSigned.ShowCheckBox = True
        Me.dtPickPipeSigned.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeSigned.TabIndex = 0
        '
        'lblDtPipeSigned
        '
        Me.lblDtPipeSigned.AutoSize = True
        Me.lblDtPipeSigned.Location = New System.Drawing.Point(104, 8)
        Me.lblDtPipeSigned.Name = "lblDtPipeSigned"
        Me.lblDtPipeSigned.Size = New System.Drawing.Size(76, 17)
        Me.lblDtPipeSigned.TabIndex = 127
        Me.lblDtPipeSigned.Text = "Date Signed:"
        '
        'Panel9
        '
        Me.Panel9.Controls.Add(Me.lblPipeInstallerOath)
        Me.Panel9.Controls.Add(Me.lblPipeInstallerOathDisplay)
        Me.Panel9.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel9.Location = New System.Drawing.Point(3, 832)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(846, 24)
        Me.Panel9.TabIndex = 14
        '
        'lblPipeInstallerOath
        '
        Me.lblPipeInstallerOath.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPipeInstallerOath.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPipeInstallerOath.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPipeInstallerOath.Location = New System.Drawing.Point(16, 0)
        Me.lblPipeInstallerOath.Name = "lblPipeInstallerOath"
        Me.lblPipeInstallerOath.Size = New System.Drawing.Size(830, 24)
        Me.lblPipeInstallerOath.TabIndex = 1
        Me.lblPipeInstallerOath.Text = "Installer Oath"
        Me.lblPipeInstallerOath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPipeInstallerOathDisplay
        '
        Me.lblPipeInstallerOathDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPipeInstallerOathDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPipeInstallerOathDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPipeInstallerOathDisplay.Name = "lblPipeInstallerOathDisplay"
        Me.lblPipeInstallerOathDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPipeInstallerOathDisplay.TabIndex = 0
        Me.lblPipeInstallerOathDisplay.Text = "-"
        Me.lblPipeInstallerOathDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPipeRelease
        '
        Me.pnlPipeRelease.Controls.Add(Me.dtElectronicDeviceInspected)
        Me.pnlPipeRelease.Controls.Add(Me.LblElectronicDeviceInspected)
        Me.pnlPipeRelease.Controls.Add(Me.grpBxReleaseDetectionGroup2)
        Me.pnlPipeRelease.Controls.Add(Me.grpBxPipeReleaseDetectionGroup1)
        Me.pnlPipeRelease.Controls.Add(Me.LblSecondaryContainmentInspected)
        Me.pnlPipeRelease.Controls.Add(Me.dtSecondaryContainmentInspected)
        Me.pnlPipeRelease.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeRelease.Location = New System.Drawing.Point(3, 648)
        Me.pnlPipeRelease.Name = "pnlPipeRelease"
        Me.pnlPipeRelease.Size = New System.Drawing.Size(846, 184)
        Me.pnlPipeRelease.TabIndex = 4
        '
        'dtElectronicDeviceInspected
        '
        Me.dtElectronicDeviceInspected.Checked = False
        Me.dtElectronicDeviceInspected.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtElectronicDeviceInspected.Location = New System.Drawing.Point(216, 160)
        Me.dtElectronicDeviceInspected.Name = "dtElectronicDeviceInspected"
        Me.dtElectronicDeviceInspected.ShowCheckBox = True
        Me.dtElectronicDeviceInspected.Size = New System.Drawing.Size(104, 21)
        Me.dtElectronicDeviceInspected.TabIndex = 139
        Me.dtElectronicDeviceInspected.Value = New Date(2008, 10, 1, 0, 0, 0, 0)
        '
        'LblElectronicDeviceInspected
        '
        Me.LblElectronicDeviceInspected.AutoSize = True
        Me.LblElectronicDeviceInspected.Location = New System.Drawing.Point(16, 160)
        Me.LblElectronicDeviceInspected.Name = "LblElectronicDeviceInspected"
        Me.LblElectronicDeviceInspected.Size = New System.Drawing.Size(190, 17)
        Me.LblElectronicDeviceInspected.TabIndex = 140
        Me.LblElectronicDeviceInspected.Text = "Electronic Device Inspected Date:"
        '
        'grpBxReleaseDetectionGroup2
        '
        Me.grpBxReleaseDetectionGroup2.Controls.Add(Me.dtPickPipeLeakDetectorTest)
        Me.grpBxReleaseDetectionGroup2.Controls.Add(Me.lblPipeReleaseDetection2)
        Me.grpBxReleaseDetectionGroup2.Controls.Add(Me.cmbPipeReleaseDetection2)
        Me.grpBxReleaseDetectionGroup2.Controls.Add(Me.lblPipeALLDTestDate)
        Me.grpBxReleaseDetectionGroup2.Location = New System.Drawing.Point(16, 79)
        Me.grpBxReleaseDetectionGroup2.Name = "grpBxReleaseDetectionGroup2"
        Me.grpBxReleaseDetectionGroup2.Size = New System.Drawing.Size(648, 72)
        Me.grpBxReleaseDetectionGroup2.TabIndex = 1
        Me.grpBxReleaseDetectionGroup2.TabStop = False
        Me.grpBxReleaseDetectionGroup2.Text = "Group 2"
        '
        'dtPickPipeLeakDetectorTest
        '
        Me.dtPickPipeLeakDetectorTest.Checked = False
        Me.dtPickPipeLeakDetectorTest.Enabled = False
        Me.dtPickPipeLeakDetectorTest.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeLeakDetectorTest.Location = New System.Drawing.Point(296, 47)
        Me.dtPickPipeLeakDetectorTest.Name = "dtPickPipeLeakDetectorTest"
        Me.dtPickPipeLeakDetectorTest.ShowCheckBox = True
        Me.dtPickPipeLeakDetectorTest.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeLeakDetectorTest.TabIndex = 1
        '
        'lblPipeReleaseDetection2
        '
        Me.lblPipeReleaseDetection2.AutoSize = True
        Me.lblPipeReleaseDetection2.Location = New System.Drawing.Point(48, 20)
        Me.lblPipeReleaseDetection2.Name = "lblPipeReleaseDetection2"
        Me.lblPipeReleaseDetection2.Size = New System.Drawing.Size(109, 17)
        Me.lblPipeReleaseDetection2.TabIndex = 133
        Me.lblPipeReleaseDetection2.Text = "Release Detection:"
        '
        'cmbPipeReleaseDetection2
        '
        Me.cmbPipeReleaseDetection2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeReleaseDetection2.DropDownWidth = 300
        Me.cmbPipeReleaseDetection2.Items.AddRange(New Object() {"Continous Intersitial Monitoring"})
        Me.cmbPipeReleaseDetection2.Location = New System.Drawing.Point(168, 20)
        Me.cmbPipeReleaseDetection2.Name = "cmbPipeReleaseDetection2"
        Me.cmbPipeReleaseDetection2.Size = New System.Drawing.Size(336, 23)
        Me.cmbPipeReleaseDetection2.TabIndex = 0
        '
        'lblPipeALLDTestDate
        '
        Me.lblPipeALLDTestDate.AutoSize = True
        Me.lblPipeALLDTestDate.Location = New System.Drawing.Point(48, 47)
        Me.lblPipeALLDTestDate.Name = "lblPipeALLDTestDate"
        Me.lblPipeALLDTestDate.Size = New System.Drawing.Size(226, 17)
        Me.lblPipeALLDTestDate.TabIndex = 134
        Me.lblPipeALLDTestDate.Text = "Automatic Line Leak Detector Test Date:"
        '
        'grpBxPipeReleaseDetectionGroup1
        '
        Me.grpBxPipeReleaseDetectionGroup1.Controls.Add(Me.dtPickPipeTightnessTest)
        Me.grpBxPipeReleaseDetectionGroup1.Controls.Add(Me.lblPipeReleaseDetection)
        Me.grpBxPipeReleaseDetectionGroup1.Controls.Add(Me.cmbPipeReleaseDetection1)
        Me.grpBxPipeReleaseDetectionGroup1.Controls.Add(Me.lblLPipeTightnessTstDt)
        Me.grpBxPipeReleaseDetectionGroup1.Location = New System.Drawing.Point(16, 7)
        Me.grpBxPipeReleaseDetectionGroup1.Name = "grpBxPipeReleaseDetectionGroup1"
        Me.grpBxPipeReleaseDetectionGroup1.Size = New System.Drawing.Size(648, 72)
        Me.grpBxPipeReleaseDetectionGroup1.TabIndex = 0
        Me.grpBxPipeReleaseDetectionGroup1.TabStop = False
        Me.grpBxPipeReleaseDetectionGroup1.Text = "Group 1"
        '
        'dtPickPipeTightnessTest
        '
        Me.dtPickPipeTightnessTest.Checked = False
        Me.dtPickPipeTightnessTest.Enabled = False
        Me.dtPickPipeTightnessTest.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeTightnessTest.Location = New System.Drawing.Point(232, 45)
        Me.dtPickPipeTightnessTest.Name = "dtPickPipeTightnessTest"
        Me.dtPickPipeTightnessTest.ShowCheckBox = True
        Me.dtPickPipeTightnessTest.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeTightnessTest.TabIndex = 1
        '
        'lblPipeReleaseDetection
        '
        Me.lblPipeReleaseDetection.AutoSize = True
        Me.lblPipeReleaseDetection.Location = New System.Drawing.Point(48, 19)
        Me.lblPipeReleaseDetection.Name = "lblPipeReleaseDetection"
        Me.lblPipeReleaseDetection.Size = New System.Drawing.Size(109, 17)
        Me.lblPipeReleaseDetection.TabIndex = 133
        Me.lblPipeReleaseDetection.Text = "Release Detection:"
        '
        'cmbPipeReleaseDetection1
        '
        Me.cmbPipeReleaseDetection1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeReleaseDetection1.DropDownWidth = 300
        Me.cmbPipeReleaseDetection1.Items.AddRange(New Object() {"Line Tightness Testing and CAP"})
        Me.cmbPipeReleaseDetection1.Location = New System.Drawing.Point(168, 16)
        Me.cmbPipeReleaseDetection1.Name = "cmbPipeReleaseDetection1"
        Me.cmbPipeReleaseDetection1.Size = New System.Drawing.Size(288, 23)
        Me.cmbPipeReleaseDetection1.TabIndex = 0
        '
        'lblLPipeTightnessTstDt
        '
        Me.lblLPipeTightnessTstDt.AutoSize = True
        Me.lblLPipeTightnessTstDt.Location = New System.Drawing.Point(48, 46)
        Me.lblLPipeTightnessTstDt.Name = "lblLPipeTightnessTstDt"
        Me.lblLPipeTightnessTstDt.Size = New System.Drawing.Size(171, 17)
        Me.lblLPipeTightnessTstDt.TabIndex = 134
        Me.lblLPipeTightnessTstDt.Text = "Last Pipe Tightness Test date:"
        '
        'LblSecondaryContainmentInspected
        '
        Me.LblSecondaryContainmentInspected.AutoSize = True
        Me.LblSecondaryContainmentInspected.Location = New System.Drawing.Point(336, 160)
        Me.LblSecondaryContainmentInspected.Name = "LblSecondaryContainmentInspected"
        Me.LblSecondaryContainmentInspected.Size = New System.Drawing.Size(223, 17)
        Me.LblSecondaryContainmentInspected.TabIndex = 160
        Me.LblSecondaryContainmentInspected.Text = "Date Secondary Containment Inspected"
        '
        'dtSecondaryContainmentInspected
        '
        Me.dtSecondaryContainmentInspected.Checked = False
        Me.dtSecondaryContainmentInspected.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtSecondaryContainmentInspected.Location = New System.Drawing.Point(560, 160)
        Me.dtSecondaryContainmentInspected.Name = "dtSecondaryContainmentInspected"
        Me.dtSecondaryContainmentInspected.ShowCheckBox = True
        Me.dtSecondaryContainmentInspected.Size = New System.Drawing.Size(104, 21)
        Me.dtSecondaryContainmentInspected.TabIndex = 159
        Me.dtSecondaryContainmentInspected.Value = New Date(2008, 10, 1, 0, 0, 0, 0)
        '
        'Panel11
        '
        Me.Panel11.Controls.Add(Me.lblPipeReleaseHead)
        Me.Panel11.Controls.Add(Me.lblPipeReleaseDisplay)
        Me.Panel11.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel11.Location = New System.Drawing.Point(3, 624)
        Me.Panel11.Name = "Panel11"
        Me.Panel11.Size = New System.Drawing.Size(846, 24)
        Me.Panel11.TabIndex = 12
        '
        'lblPipeReleaseHead
        '
        Me.lblPipeReleaseHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPipeReleaseHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPipeReleaseHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPipeReleaseHead.Location = New System.Drawing.Point(16, 0)
        Me.lblPipeReleaseHead.Name = "lblPipeReleaseHead"
        Me.lblPipeReleaseHead.Size = New System.Drawing.Size(830, 24)
        Me.lblPipeReleaseHead.TabIndex = 1
        Me.lblPipeReleaseHead.Text = "Release Detection"
        Me.lblPipeReleaseHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPipeReleaseDisplay
        '
        Me.lblPipeReleaseDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPipeReleaseDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPipeReleaseDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPipeReleaseDisplay.Name = "lblPipeReleaseDisplay"
        Me.lblPipeReleaseDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPipeReleaseDisplay.TabIndex = 0
        Me.lblPipeReleaseDisplay.Text = "-"
        Me.lblPipeReleaseDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPipeType
        '
        Me.pnlPipeType.Controls.Add(Me.dtSheerValueTest)
        Me.pnlPipeType.Controls.Add(Me.LblSheerValueTestDate)
        Me.pnlPipeType.Controls.Add(Me.cmbPipeType)
        Me.pnlPipeType.Controls.Add(Me.lblPipeCapacity)
        Me.pnlPipeType.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeType.Location = New System.Drawing.Point(3, 560)
        Me.pnlPipeType.Name = "pnlPipeType"
        Me.pnlPipeType.Size = New System.Drawing.Size(846, 64)
        Me.pnlPipeType.TabIndex = 3
        '
        'dtSheerValueTest
        '
        Me.dtSheerValueTest.Checked = False
        Me.dtSheerValueTest.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtSheerValueTest.Location = New System.Drawing.Point(184, 40)
        Me.dtSheerValueTest.Name = "dtSheerValueTest"
        Me.dtSheerValueTest.ShowCheckBox = True
        Me.dtSheerValueTest.Size = New System.Drawing.Size(104, 21)
        Me.dtSheerValueTest.TabIndex = 157
        Me.dtSheerValueTest.Value = New Date(2008, 10, 1, 0, 0, 0, 0)
        '
        'LblSheerValueTestDate
        '
        Me.LblSheerValueTestDate.AutoSize = True
        Me.LblSheerValueTestDate.Location = New System.Drawing.Point(40, 40)
        Me.LblSheerValueTestDate.Name = "LblSheerValueTestDate"
        Me.LblSheerValueTestDate.Size = New System.Drawing.Size(133, 17)
        Me.LblSheerValueTestDate.TabIndex = 158
        Me.LblSheerValueTestDate.Text = "Shear Value Test Date:"
        '
        'cmbPipeType
        '
        Me.cmbPipeType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeType.DropDownWidth = 220
        Me.cmbPipeType.Location = New System.Drawing.Point(184, 8)
        Me.cmbPipeType.Name = "cmbPipeType"
        Me.cmbPipeType.Size = New System.Drawing.Size(256, 23)
        Me.cmbPipeType.TabIndex = 0
        '
        'lblPipeCapacity
        '
        Me.lblPipeCapacity.AutoSize = True
        Me.lblPipeCapacity.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeCapacity.Location = New System.Drawing.Point(69, 8)
        Me.lblPipeCapacity.Name = "lblPipeCapacity"
        Me.lblPipeCapacity.Size = New System.Drawing.Size(64, 17)
        Me.lblPipeCapacity.TabIndex = 156
        Me.lblPipeCapacity.Text = "Pipe Type:"
        '
        'Panel15
        '
        Me.Panel15.Controls.Add(Me.lblPipeTypeCaption)
        Me.Panel15.Controls.Add(Me.lblPipeType)
        Me.Panel15.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel15.Location = New System.Drawing.Point(3, 536)
        Me.Panel15.Name = "Panel15"
        Me.Panel15.Size = New System.Drawing.Size(846, 24)
        Me.Panel15.TabIndex = 18
        '
        'lblPipeTypeCaption
        '
        Me.lblPipeTypeCaption.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPipeTypeCaption.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPipeTypeCaption.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPipeTypeCaption.Location = New System.Drawing.Point(16, 0)
        Me.lblPipeTypeCaption.Name = "lblPipeTypeCaption"
        Me.lblPipeTypeCaption.Size = New System.Drawing.Size(830, 24)
        Me.lblPipeTypeCaption.TabIndex = 5
        Me.lblPipeTypeCaption.Text = "Piping(Type)"
        Me.lblPipeTypeCaption.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPipeType
        '
        Me.lblPipeType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPipeType.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPipeType.Location = New System.Drawing.Point(0, 0)
        Me.lblPipeType.Name = "lblPipeType"
        Me.lblPipeType.Size = New System.Drawing.Size(16, 24)
        Me.lblPipeType.TabIndex = 4
        Me.lblPipeType.Text = "-"
        Me.lblPipeType.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPipeMaterial
        '
        Me.pnlPipeMaterial.Controls.Add(Me.lblEmergencyPowerPipe)
        Me.pnlPipeMaterial.Controls.Add(Me.chkEmergencyPowerPipe)
        Me.pnlPipeMaterial.Controls.Add(Me.lblPipeManufacturer)
        Me.pnlPipeMaterial.Controls.Add(Me.cmbPipeManufacturerID)
        Me.pnlPipeMaterial.Controls.Add(Me.dtPickPipeTerminationCPLastTested)
        Me.pnlPipeMaterial.Controls.Add(Me.lblPipeTerminationCPLastTest)
        Me.pnlPipeMaterial.Controls.Add(Me.dtPickPipeTerminationCPInstalled)
        Me.pnlPipeMaterial.Controls.Add(Me.lblPipeTerminationCPInstalled)
        Me.pnlPipeMaterial.Controls.Add(Me.grpBxPipeTerminationAtDispenser)
        Me.pnlPipeMaterial.Controls.Add(Me.grpBxPipeTerminationAtTank)
        Me.pnlPipeMaterial.Controls.Add(Me.grpBxPipeContainmentSumpsLocation)
        Me.pnlPipeMaterial.Controls.Add(Me.dtPickPipeCPLastTest)
        Me.pnlPipeMaterial.Controls.Add(Me.dtPickPipeCPInstalled)
        Me.pnlPipeMaterial.Controls.Add(Me.lblPipeMaterial)
        Me.pnlPipeMaterial.Controls.Add(Me.lblPipeOption)
        Me.pnlPipeMaterial.Controls.Add(Me.cmbPipeMaterial)
        Me.pnlPipeMaterial.Controls.Add(Me.cmbPipeOptions)
        Me.pnlPipeMaterial.Controls.Add(Me.lblPipeCPType)
        Me.pnlPipeMaterial.Controls.Add(Me.cmbPipeCPType)
        Me.pnlPipeMaterial.Controls.Add(Me.lblDatePipeCPInstalled)
        Me.pnlPipeMaterial.Controls.Add(Me.lblDtPipeLastTested)
        Me.pnlPipeMaterial.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeMaterial.Location = New System.Drawing.Point(3, 224)
        Me.pnlPipeMaterial.Name = "pnlPipeMaterial"
        Me.pnlPipeMaterial.Size = New System.Drawing.Size(846, 312)
        Me.pnlPipeMaterial.TabIndex = 2
        '
        'lblEmergencyPowerPipe
        '
        Me.lblEmergencyPowerPipe.Location = New System.Drawing.Point(446, 8)
        Me.lblEmergencyPowerPipe.Name = "lblEmergencyPowerPipe"
        Me.lblEmergencyPowerPipe.Size = New System.Drawing.Size(184, 23)
        Me.lblEmergencyPowerPipe.TabIndex = 171
        Me.lblEmergencyPowerPipe.Text = "Tank used for emergency power"
        '
        'chkEmergencyPowerPipe
        '
        Me.chkEmergencyPowerPipe.Enabled = False
        Me.chkEmergencyPowerPipe.Location = New System.Drawing.Point(630, 7)
        Me.chkEmergencyPowerPipe.Name = "chkEmergencyPowerPipe"
        Me.chkEmergencyPowerPipe.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkEmergencyPowerPipe.Size = New System.Drawing.Size(16, 24)
        Me.chkEmergencyPowerPipe.TabIndex = 170
        Me.chkEmergencyPowerPipe.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPipeManufacturer
        '
        Me.lblPipeManufacturer.AutoSize = True
        Me.lblPipeManufacturer.Location = New System.Drawing.Point(56, 82)
        Me.lblPipeManufacturer.Name = "lblPipeManufacturer"
        Me.lblPipeManufacturer.Size = New System.Drawing.Size(81, 17)
        Me.lblPipeManufacturer.TabIndex = 169
        Me.lblPipeManufacturer.Text = "Manufacturer:"
        Me.lblPipeManufacturer.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbPipeManufacturerID
        '
        Me.cmbPipeManufacturerID.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeManufacturerID.DropDownWidth = 300
        Me.cmbPipeManufacturerID.Location = New System.Drawing.Point(152, 82)
        Me.cmbPipeManufacturerID.Name = "cmbPipeManufacturerID"
        Me.cmbPipeManufacturerID.Size = New System.Drawing.Size(248, 23)
        Me.cmbPipeManufacturerID.TabIndex = 3
        '
        'dtPickPipeTerminationCPLastTested
        '
        Me.dtPickPipeTerminationCPLastTested.Checked = False
        Me.dtPickPipeTerminationCPLastTested.Enabled = False
        Me.dtPickPipeTerminationCPLastTested.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeTerminationCPLastTested.Location = New System.Drawing.Point(536, 280)
        Me.dtPickPipeTerminationCPLastTested.Name = "dtPickPipeTerminationCPLastTested"
        Me.dtPickPipeTerminationCPLastTested.ShowCheckBox = True
        Me.dtPickPipeTerminationCPLastTested.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeTerminationCPLastTested.TabIndex = 11
        '
        'lblPipeTerminationCPLastTest
        '
        Me.lblPipeTerminationCPLastTest.AutoSize = True
        Me.lblPipeTerminationCPLastTest.Location = New System.Drawing.Point(352, 280)
        Me.lblPipeTerminationCPLastTest.Name = "lblPipeTerminationCPLastTest"
        Me.lblPipeTerminationCPLastTest.Size = New System.Drawing.Size(185, 17)
        Me.lblPipeTerminationCPLastTest.TabIndex = 165
        Me.lblPipeTerminationCPLastTest.Text = "Termination CP Last Tested  On:"
        Me.lblPipeTerminationCPLastTest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtPickPipeTerminationCPInstalled
        '
        Me.dtPickPipeTerminationCPInstalled.Checked = False
        Me.dtPickPipeTerminationCPInstalled.Enabled = False
        Me.dtPickPipeTerminationCPInstalled.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeTerminationCPInstalled.Location = New System.Drawing.Point(184, 280)
        Me.dtPickPipeTerminationCPInstalled.Name = "dtPickPipeTerminationCPInstalled"
        Me.dtPickPipeTerminationCPInstalled.ShowCheckBox = True
        Me.dtPickPipeTerminationCPInstalled.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeTerminationCPInstalled.TabIndex = 10
        '
        'lblPipeTerminationCPInstalled
        '
        Me.lblPipeTerminationCPInstalled.AutoSize = True
        Me.lblPipeTerminationCPInstalled.Location = New System.Drawing.Point(16, 280)
        Me.lblPipeTerminationCPInstalled.Name = "lblPipeTerminationCPInstalled"
        Me.lblPipeTerminationCPInstalled.Size = New System.Drawing.Size(166, 17)
        Me.lblPipeTerminationCPInstalled.TabIndex = 163
        Me.lblPipeTerminationCPInstalled.Text = "Termination CP Installed  On:"
        Me.lblPipeTerminationCPInstalled.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpBxPipeTerminationAtDispenser
        '
        Me.grpBxPipeTerminationAtDispenser.Controls.Add(Me.cmbPipeTerminationDispenserCPType)
        Me.grpBxPipeTerminationAtDispenser.Controls.Add(Me.cmbPipeTerminationDispenserType)
        Me.grpBxPipeTerminationAtDispenser.Controls.Add(Me.lblPipeTerminationDispenserCPType)
        Me.grpBxPipeTerminationAtDispenser.Controls.Add(Me.lblPipeTerminationDispenserType)
        Me.grpBxPipeTerminationAtDispenser.Location = New System.Drawing.Point(8, 168)
        Me.grpBxPipeTerminationAtDispenser.Name = "grpBxPipeTerminationAtDispenser"
        Me.grpBxPipeTerminationAtDispenser.Size = New System.Drawing.Size(648, 48)
        Me.grpBxPipeTerminationAtDispenser.TabIndex = 8
        Me.grpBxPipeTerminationAtDispenser.TabStop = False
        Me.grpBxPipeTerminationAtDispenser.Text = "Termination at Dispenser"
        '
        'cmbPipeTerminationDispenserCPType
        '
        Me.cmbPipeTerminationDispenserCPType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeTerminationDispenserCPType.DropDownWidth = 300
        Me.cmbPipeTerminationDispenserCPType.Location = New System.Drawing.Point(424, 16)
        Me.cmbPipeTerminationDispenserCPType.Name = "cmbPipeTerminationDispenserCPType"
        Me.cmbPipeTerminationDispenserCPType.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeTerminationDispenserCPType.TabIndex = 1
        '
        'cmbPipeTerminationDispenserType
        '
        Me.cmbPipeTerminationDispenserType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeTerminationDispenserType.DropDownWidth = 300
        Me.cmbPipeTerminationDispenserType.Items.AddRange(New Object() {"Coated Wrapped/Cathodically Protected", "Coated Wrapped/Cathodically Protected and CAP"})
        Me.cmbPipeTerminationDispenserType.Location = New System.Drawing.Point(112, 17)
        Me.cmbPipeTerminationDispenserType.Name = "cmbPipeTerminationDispenserType"
        Me.cmbPipeTerminationDispenserType.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeTerminationDispenserType.TabIndex = 0
        '
        'lblPipeTerminationDispenserCPType
        '
        Me.lblPipeTerminationDispenserCPType.AutoSize = True
        Me.lblPipeTerminationDispenserCPType.Location = New System.Drawing.Point(360, 19)
        Me.lblPipeTerminationDispenserCPType.Name = "lblPipeTerminationDispenserCPType"
        Me.lblPipeTerminationDispenserCPType.Size = New System.Drawing.Size(56, 17)
        Me.lblPipeTerminationDispenserCPType.TabIndex = 157
        Me.lblPipeTerminationDispenserCPType.Text = "CP Type:"
        Me.lblPipeTerminationDispenserCPType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPipeTerminationDispenserType
        '
        Me.lblPipeTerminationDispenserType.AutoSize = True
        Me.lblPipeTerminationDispenserType.Location = New System.Drawing.Point(72, 20)
        Me.lblPipeTerminationDispenserType.Name = "lblPipeTerminationDispenserType"
        Me.lblPipeTerminationDispenserType.Size = New System.Drawing.Size(35, 17)
        Me.lblPipeTerminationDispenserType.TabIndex = 155
        Me.lblPipeTerminationDispenserType.Text = "Type:"
        Me.lblPipeTerminationDispenserType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpBxPipeTerminationAtTank
        '
        Me.grpBxPipeTerminationAtTank.Controls.Add(Me.cmbPipeTerminationTankCPType)
        Me.grpBxPipeTerminationAtTank.Controls.Add(Me.cmbPipeTerminationTankType)
        Me.grpBxPipeTerminationAtTank.Controls.Add(Me.lblPipeTerminationTankCPType)
        Me.grpBxPipeTerminationAtTank.Controls.Add(Me.lblPipeTerminationTankType)
        Me.grpBxPipeTerminationAtTank.Location = New System.Drawing.Point(7, 224)
        Me.grpBxPipeTerminationAtTank.Name = "grpBxPipeTerminationAtTank"
        Me.grpBxPipeTerminationAtTank.Size = New System.Drawing.Size(648, 48)
        Me.grpBxPipeTerminationAtTank.TabIndex = 9
        Me.grpBxPipeTerminationAtTank.TabStop = False
        Me.grpBxPipeTerminationAtTank.Text = "Termination at Tank"
        '
        'cmbPipeTerminationTankCPType
        '
        Me.cmbPipeTerminationTankCPType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeTerminationTankCPType.DropDownWidth = 300
        Me.cmbPipeTerminationTankCPType.Location = New System.Drawing.Point(424, 13)
        Me.cmbPipeTerminationTankCPType.Name = "cmbPipeTerminationTankCPType"
        Me.cmbPipeTerminationTankCPType.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeTerminationTankCPType.TabIndex = 1
        '
        'cmbPipeTerminationTankType
        '
        Me.cmbPipeTerminationTankType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeTerminationTankType.DropDownWidth = 300
        Me.cmbPipeTerminationTankType.Items.AddRange(New Object() {"Coated Wrapped/Cathodically Protected", "Coated Wrapped/Cathodically Protected and CAP"})
        Me.cmbPipeTerminationTankType.Location = New System.Drawing.Point(112, 13)
        Me.cmbPipeTerminationTankType.Name = "cmbPipeTerminationTankType"
        Me.cmbPipeTerminationTankType.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeTerminationTankType.TabIndex = 0
        '
        'lblPipeTerminationTankCPType
        '
        Me.lblPipeTerminationTankCPType.AutoSize = True
        Me.lblPipeTerminationTankCPType.Location = New System.Drawing.Point(360, 16)
        Me.lblPipeTerminationTankCPType.Name = "lblPipeTerminationTankCPType"
        Me.lblPipeTerminationTankCPType.Size = New System.Drawing.Size(56, 17)
        Me.lblPipeTerminationTankCPType.TabIndex = 162
        Me.lblPipeTerminationTankCPType.Text = "CP Type:"
        Me.lblPipeTerminationTankCPType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPipeTerminationTankType
        '
        Me.lblPipeTerminationTankType.AutoSize = True
        Me.lblPipeTerminationTankType.Location = New System.Drawing.Point(72, 17)
        Me.lblPipeTerminationTankType.Name = "lblPipeTerminationTankType"
        Me.lblPipeTerminationTankType.Size = New System.Drawing.Size(35, 17)
        Me.lblPipeTerminationTankType.TabIndex = 161
        Me.lblPipeTerminationTankType.Text = "Type:"
        Me.lblPipeTerminationTankType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'grpBxPipeContainmentSumpsLocation
        '
        Me.grpBxPipeContainmentSumpsLocation.Controls.Add(Me.chkPipeSumpsAtTank)
        Me.grpBxPipeContainmentSumpsLocation.Controls.Add(Me.chkPipeSumpsAtDispenser)
        Me.grpBxPipeContainmentSumpsLocation.Location = New System.Drawing.Point(8, 112)
        Me.grpBxPipeContainmentSumpsLocation.Name = "grpBxPipeContainmentSumpsLocation"
        Me.grpBxPipeContainmentSumpsLocation.Size = New System.Drawing.Size(648, 48)
        Me.grpBxPipeContainmentSumpsLocation.TabIndex = 7
        Me.grpBxPipeContainmentSumpsLocation.TabStop = False
        Me.grpBxPipeContainmentSumpsLocation.Text = "Containment Sumps Located At"
        '
        'chkPipeSumpsAtTank
        '
        Me.chkPipeSumpsAtTank.Location = New System.Drawing.Point(487, 19)
        Me.chkPipeSumpsAtTank.Name = "chkPipeSumpsAtTank"
        Me.chkPipeSumpsAtTank.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkPipeSumpsAtTank.Size = New System.Drawing.Size(56, 24)
        Me.chkPipeSumpsAtTank.TabIndex = 1
        Me.chkPipeSumpsAtTank.Text = "Tank"
        Me.chkPipeSumpsAtTank.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'chkPipeSumpsAtDispenser
        '
        Me.chkPipeSumpsAtDispenser.Location = New System.Drawing.Point(112, 20)
        Me.chkPipeSumpsAtDispenser.Name = "chkPipeSumpsAtDispenser"
        Me.chkPipeSumpsAtDispenser.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkPipeSumpsAtDispenser.Size = New System.Drawing.Size(80, 24)
        Me.chkPipeSumpsAtDispenser.TabIndex = 0
        Me.chkPipeSumpsAtDispenser.Text = "Dispenser"
        Me.chkPipeSumpsAtDispenser.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dtPickPipeCPLastTest
        '
        Me.dtPickPipeCPLastTest.Checked = False
        Me.dtPickPipeCPLastTest.Enabled = False
        Me.dtPickPipeCPLastTest.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeCPLastTest.Location = New System.Drawing.Point(544, 59)
        Me.dtPickPipeCPLastTest.Name = "dtPickPipeCPLastTest"
        Me.dtPickPipeCPLastTest.ShowCheckBox = True
        Me.dtPickPipeCPLastTest.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeCPLastTest.TabIndex = 6
        '
        'dtPickPipeCPInstalled
        '
        Me.dtPickPipeCPInstalled.Checked = False
        Me.dtPickPipeCPInstalled.Enabled = False
        Me.dtPickPipeCPInstalled.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeCPInstalled.Location = New System.Drawing.Point(544, 35)
        Me.dtPickPipeCPInstalled.Name = "dtPickPipeCPInstalled"
        Me.dtPickPipeCPInstalled.ShowCheckBox = True
        Me.dtPickPipeCPInstalled.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeCPInstalled.TabIndex = 5
        '
        'lblPipeMaterial
        '
        Me.lblPipeMaterial.AutoSize = True
        Me.lblPipeMaterial.Location = New System.Drawing.Point(64, 8)
        Me.lblPipeMaterial.Name = "lblPipeMaterial"
        Me.lblPipeMaterial.Size = New System.Drawing.Size(80, 17)
        Me.lblPipeMaterial.TabIndex = 144
        Me.lblPipeMaterial.Text = "Pipe Material:"
        Me.lblPipeMaterial.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPipeOption
        '
        Me.lblPipeOption.AutoSize = True
        Me.lblPipeOption.Location = New System.Drawing.Point(8, 33)
        Me.lblPipeOption.Name = "lblPipeOption"
        Me.lblPipeOption.Size = New System.Drawing.Size(135, 17)
        Me.lblPipeOption.TabIndex = 145
        Me.lblPipeOption.Text = "Secondary Pipe Option:"
        Me.lblPipeOption.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbPipeMaterial
        '
        Me.cmbPipeMaterial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeMaterial.DropDownWidth = 300
        Me.cmbPipeMaterial.Location = New System.Drawing.Point(152, 8)
        Me.cmbPipeMaterial.Name = "cmbPipeMaterial"
        Me.cmbPipeMaterial.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeMaterial.TabIndex = 0
        '
        'cmbPipeOptions
        '
        Me.cmbPipeOptions.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeOptions.DropDownWidth = 300
        Me.cmbPipeOptions.Location = New System.Drawing.Point(152, 33)
        Me.cmbPipeOptions.Name = "cmbPipeOptions"
        Me.cmbPipeOptions.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeOptions.TabIndex = 1
        '
        'lblPipeCPType
        '
        Me.lblPipeCPType.AutoSize = True
        Me.lblPipeCPType.Location = New System.Drawing.Point(56, 57)
        Me.lblPipeCPType.Name = "lblPipeCPType"
        Me.lblPipeCPType.Size = New System.Drawing.Size(84, 17)
        Me.lblPipeCPType.TabIndex = 146
        Me.lblPipeCPType.Text = "Pipe CP Type:"
        Me.lblPipeCPType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbPipeCPType
        '
        Me.cmbPipeCPType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeCPType.DropDownWidth = 300
        Me.cmbPipeCPType.Location = New System.Drawing.Point(152, 57)
        Me.cmbPipeCPType.Name = "cmbPipeCPType"
        Me.cmbPipeCPType.Size = New System.Drawing.Size(208, 23)
        Me.cmbPipeCPType.TabIndex = 2
        '
        'lblDatePipeCPInstalled
        '
        Me.lblDatePipeCPInstalled.AutoSize = True
        Me.lblDatePipeCPInstalled.Location = New System.Drawing.Point(432, 35)
        Me.lblDatePipeCPInstalled.Name = "lblDatePipeCPInstalled"
        Me.lblDatePipeCPInstalled.Size = New System.Drawing.Size(104, 17)
        Me.lblDatePipeCPInstalled.TabIndex = 147
        Me.lblDatePipeCPInstalled.Text = "Date CP Installed:"
        Me.lblDatePipeCPInstalled.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDtPipeLastTested
        '
        Me.lblDtPipeLastTested.AutoSize = True
        Me.lblDtPipeLastTested.Location = New System.Drawing.Point(392, 59)
        Me.lblDtPipeLastTested.Name = "lblDtPipeLastTested"
        Me.lblDtPipeLastTested.Size = New System.Drawing.Size(150, 17)
        Me.lblDtPipeLastTested.TabIndex = 149
        Me.lblDtPipeLastTested.Text = "Date Pipe CP Last Tested:"
        Me.lblDtPipeLastTested.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlPipeMaterialHead
        '
        Me.pnlPipeMaterialHead.Controls.Add(Me.lblPipeMaterialHead)
        Me.pnlPipeMaterialHead.Controls.Add(Me.lblPipeMaterialDisplay)
        Me.pnlPipeMaterialHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeMaterialHead.Location = New System.Drawing.Point(3, 200)
        Me.pnlPipeMaterialHead.Name = "pnlPipeMaterialHead"
        Me.pnlPipeMaterialHead.Size = New System.Drawing.Size(846, 24)
        Me.pnlPipeMaterialHead.TabIndex = 9
        '
        'lblPipeMaterialHead
        '
        Me.lblPipeMaterialHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPipeMaterialHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPipeMaterialHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPipeMaterialHead.Location = New System.Drawing.Point(16, 0)
        Me.lblPipeMaterialHead.Name = "lblPipeMaterialHead"
        Me.lblPipeMaterialHead.Size = New System.Drawing.Size(830, 24)
        Me.lblPipeMaterialHead.TabIndex = 1
        Me.lblPipeMaterialHead.Text = "Pipe Material of Construction"
        Me.lblPipeMaterialHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPipeMaterialDisplay
        '
        Me.lblPipeMaterialDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPipeMaterialDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPipeMaterialDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPipeMaterialDisplay.Name = "lblPipeMaterialDisplay"
        Me.lblPipeMaterialDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPipeMaterialDisplay.TabIndex = 0
        Me.lblPipeMaterialDisplay.Text = "-"
        Me.lblPipeMaterialDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPipeDateOfInstallation
        '
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipeFuelTypeValue)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipeCerclaNoValue)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipeCerclaNo)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipeSubstance)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipeSubstanceValue)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.dtPickDatePipePlacedInService)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipePlacedInServiceOn)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipeInstallationPlanedFor)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.dtPickPipePlannedInstallation)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipeInstalledOn)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.dtPickPipeInstalled)
        Me.pnlPipeDateOfInstallation.Controls.Add(Me.lblPipeFuelType)
        Me.pnlPipeDateOfInstallation.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeDateOfInstallation.Location = New System.Drawing.Point(3, 88)
        Me.pnlPipeDateOfInstallation.Name = "pnlPipeDateOfInstallation"
        Me.pnlPipeDateOfInstallation.Size = New System.Drawing.Size(846, 112)
        Me.pnlPipeDateOfInstallation.TabIndex = 1
        '
        'lblPipeFuelTypeValue
        '
        Me.lblPipeFuelTypeValue.Location = New System.Drawing.Point(496, 8)
        Me.lblPipeFuelTypeValue.Name = "lblPipeFuelTypeValue"
        Me.lblPipeFuelTypeValue.Size = New System.Drawing.Size(160, 23)
        Me.lblPipeFuelTypeValue.TabIndex = 176
        '
        'lblPipeCerclaNoValue
        '
        Me.lblPipeCerclaNoValue.Location = New System.Drawing.Point(176, 80)
        Me.lblPipeCerclaNoValue.Name = "lblPipeCerclaNoValue"
        Me.lblPipeCerclaNoValue.Size = New System.Drawing.Size(288, 23)
        Me.lblPipeCerclaNoValue.TabIndex = 175
        '
        'lblPipeCerclaNo
        '
        Me.lblPipeCerclaNo.Location = New System.Drawing.Point(91, 80)
        Me.lblPipeCerclaNo.Name = "lblPipeCerclaNo"
        Me.lblPipeCerclaNo.Size = New System.Drawing.Size(80, 23)
        Me.lblPipeCerclaNo.TabIndex = 174
        Me.lblPipeCerclaNo.Text = "CERCLA No:"
        '
        'lblPipeSubstance
        '
        Me.lblPipeSubstance.Location = New System.Drawing.Point(102, 56)
        Me.lblPipeSubstance.Name = "lblPipeSubstance"
        Me.lblPipeSubstance.Size = New System.Drawing.Size(72, 23)
        Me.lblPipeSubstance.TabIndex = 173
        Me.lblPipeSubstance.Text = "Substance:"
        '
        'lblPipeSubstanceValue
        '
        Me.lblPipeSubstanceValue.Location = New System.Drawing.Point(176, 56)
        Me.lblPipeSubstanceValue.Name = "lblPipeSubstanceValue"
        Me.lblPipeSubstanceValue.Size = New System.Drawing.Size(288, 23)
        Me.lblPipeSubstanceValue.TabIndex = 172
        '
        'dtPickDatePipePlacedInService
        '
        Me.dtPickDatePipePlacedInService.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickDatePipePlacedInService.Checked = False
        Me.dtPickDatePipePlacedInService.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickDatePipePlacedInService.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickDatePipePlacedInService.Location = New System.Drawing.Point(176, 32)
        Me.dtPickDatePipePlacedInService.Name = "dtPickDatePipePlacedInService"
        Me.dtPickDatePipePlacedInService.ShowCheckBox = True
        Me.dtPickDatePipePlacedInService.Size = New System.Drawing.Size(104, 21)
        Me.dtPickDatePipePlacedInService.TabIndex = 1
        Me.dtPickDatePipePlacedInService.Value = New Date(2004, 5, 27, 14, 25, 14, 234)
        '
        'lblPipePlacedInServiceOn
        '
        Me.lblPipePlacedInServiceOn.AutoSize = True
        Me.lblPipePlacedInServiceOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipePlacedInServiceOn.Location = New System.Drawing.Point(19, 32)
        Me.lblPipePlacedInServiceOn.Name = "lblPipePlacedInServiceOn"
        Me.lblPipePlacedInServiceOn.Size = New System.Drawing.Size(154, 17)
        Me.lblPipePlacedInServiceOn.TabIndex = 170
        Me.lblPipePlacedInServiceOn.Text = "Pipe Placed in Service On: "
        '
        'lblPipeInstallationPlanedFor
        '
        Me.lblPipeInstallationPlanedFor.AutoSize = True
        Me.lblPipeInstallationPlanedFor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeInstallationPlanedFor.Location = New System.Drawing.Point(576, 64)
        Me.lblPipeInstallationPlanedFor.Name = "lblPipeInstallationPlanedFor"
        Me.lblPipeInstallationPlanedFor.Size = New System.Drawing.Size(166, 17)
        Me.lblPipeInstallationPlanedFor.TabIndex = 140
        Me.lblPipeInstallationPlanedFor.Text = "Pipe Installation Planned For:"
        Me.lblPipeInstallationPlanedFor.Visible = False
        '
        'dtPickPipePlannedInstallation
        '
        Me.dtPickPipePlannedInstallation.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickPipePlannedInstallation.Checked = False
        Me.dtPickPipePlannedInstallation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickPipePlannedInstallation.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipePlannedInstallation.Location = New System.Drawing.Point(632, 64)
        Me.dtPickPipePlannedInstallation.Name = "dtPickPipePlannedInstallation"
        Me.dtPickPipePlannedInstallation.ShowCheckBox = True
        Me.dtPickPipePlannedInstallation.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipePlannedInstallation.TabIndex = 170
        Me.dtPickPipePlannedInstallation.Value = New Date(2004, 8, 7, 0, 0, 0, 0)
        Me.dtPickPipePlannedInstallation.Visible = False
        '
        'lblPipeInstalledOn
        '
        Me.lblPipeInstalledOn.AutoSize = True
        Me.lblPipeInstalledOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeInstalledOn.Location = New System.Drawing.Point(64, 8)
        Me.lblPipeInstalledOn.Name = "lblPipeInstalledOn"
        Me.lblPipeInstalledOn.Size = New System.Drawing.Size(105, 17)
        Me.lblPipeInstalledOn.TabIndex = 139
        Me.lblPipeInstalledOn.Text = " Pipe Installed On:"
        '
        'dtPickPipeInstalled
        '
        Me.dtPickPipeInstalled.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickPipeInstalled.Checked = False
        Me.dtPickPipeInstalled.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickPipeInstalled.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPipeInstalled.Location = New System.Drawing.Point(176, 8)
        Me.dtPickPipeInstalled.Name = "dtPickPipeInstalled"
        Me.dtPickPipeInstalled.ShowCheckBox = True
        Me.dtPickPipeInstalled.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPipeInstalled.TabIndex = 0
        Me.dtPickPipeInstalled.Value = New Date(2004, 5, 27, 14, 25, 14, 174)
        '
        'lblPipeFuelType
        '
        Me.lblPipeFuelType.AutoSize = True
        Me.lblPipeFuelType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeFuelType.Location = New System.Drawing.Point(424, 8)
        Me.lblPipeFuelType.Name = "lblPipeFuelType"
        Me.lblPipeFuelType.Size = New System.Drawing.Size(63, 17)
        Me.lblPipeFuelType.TabIndex = 162
        Me.lblPipeFuelType.Text = "Fuel Type:"
        '
        'Panel13
        '
        Me.Panel13.Controls.Add(Me.lblPipeDateOfInstallationCaption)
        Me.Panel13.Controls.Add(Me.lblPipeDateOfInstallation)
        Me.Panel13.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel13.Location = New System.Drawing.Point(3, 64)
        Me.Panel13.Name = "Panel13"
        Me.Panel13.Size = New System.Drawing.Size(846, 24)
        Me.Panel13.TabIndex = 16
        '
        'lblPipeDateOfInstallationCaption
        '
        Me.lblPipeDateOfInstallationCaption.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPipeDateOfInstallationCaption.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPipeDateOfInstallationCaption.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPipeDateOfInstallationCaption.Location = New System.Drawing.Point(16, 0)
        Me.lblPipeDateOfInstallationCaption.Name = "lblPipeDateOfInstallationCaption"
        Me.lblPipeDateOfInstallationCaption.Size = New System.Drawing.Size(830, 24)
        Me.lblPipeDateOfInstallationCaption.TabIndex = 3
        Me.lblPipeDateOfInstallationCaption.Text = "Date of Installation(month/year)"
        Me.lblPipeDateOfInstallationCaption.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPipeDateOfInstallation
        '
        Me.lblPipeDateOfInstallation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPipeDateOfInstallation.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPipeDateOfInstallation.Location = New System.Drawing.Point(0, 0)
        Me.lblPipeDateOfInstallation.Name = "lblPipeDateOfInstallation"
        Me.lblPipeDateOfInstallation.Size = New System.Drawing.Size(16, 24)
        Me.lblPipeDateOfInstallation.TabIndex = 2
        Me.lblPipeDateOfInstallation.Text = "-"
        Me.lblPipeDateOfInstallation.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPipeDescription
        '
        Me.pnlPipeDescription.AutoScroll = True
        Me.pnlPipeDescription.Controls.Add(Me.cmbPipeStatus)
        Me.pnlPipeDescription.Controls.Add(Me.lblPipeStatus)
        Me.pnlPipeDescription.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlPipeDescription.Location = New System.Drawing.Point(3, 27)
        Me.pnlPipeDescription.Name = "pnlPipeDescription"
        Me.pnlPipeDescription.Size = New System.Drawing.Size(846, 37)
        Me.pnlPipeDescription.TabIndex = 0
        '
        'cmbPipeStatus
        '
        Me.cmbPipeStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPipeStatus.DropDownWidth = 300
        Me.cmbPipeStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPipeStatus.ItemHeight = 15
        Me.cmbPipeStatus.Location = New System.Drawing.Point(168, 8)
        Me.cmbPipeStatus.Name = "cmbPipeStatus"
        Me.cmbPipeStatus.Size = New System.Drawing.Size(272, 23)
        Me.cmbPipeStatus.TabIndex = 0
        '
        'lblPipeStatus
        '
        Me.lblPipeStatus.AutoSize = True
        Me.lblPipeStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeStatus.Location = New System.Drawing.Point(96, 8)
        Me.lblPipeStatus.Name = "lblPipeStatus"
        Me.lblPipeStatus.Size = New System.Drawing.Size(71, 17)
        Me.lblPipeStatus.TabIndex = 137
        Me.lblPipeStatus.Text = "Pipe Status:"
        '
        'pnlPipeDescHead
        '
        Me.pnlPipeDescHead.Controls.Add(Me.lblPipeDescHead)
        Me.pnlPipeDescHead.Controls.Add(Me.lblPipeDescDisplay)
        Me.pnlPipeDescHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeDescHead.Location = New System.Drawing.Point(3, 3)
        Me.pnlPipeDescHead.Name = "pnlPipeDescHead"
        Me.pnlPipeDescHead.Size = New System.Drawing.Size(846, 24)
        Me.pnlPipeDescHead.TabIndex = 7
        '
        'lblPipeDescHead
        '
        Me.lblPipeDescHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPipeDescHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPipeDescHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPipeDescHead.Location = New System.Drawing.Point(16, 0)
        Me.lblPipeDescHead.Name = "lblPipeDescHead"
        Me.lblPipeDescHead.Size = New System.Drawing.Size(830, 24)
        Me.lblPipeDescHead.TabIndex = 1
        Me.lblPipeDescHead.Text = "Pipe Status"
        Me.lblPipeDescHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPipeDescDisplay
        '
        Me.lblPipeDescDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPipeDescDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPipeDescDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPipeDescDisplay.Name = "lblPipeDescDisplay"
        Me.lblPipeDescDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPipeDescDisplay.TabIndex = 0
        Me.lblPipeDescDisplay.Text = "-"
        Me.lblPipeDescDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPipeButtons
        '
        Me.pnlPipeButtons.Controls.Add(Me.btnPipeCancel)
        Me.pnlPipeButtons.Controls.Add(Me.btnPipeComments)
        Me.pnlPipeButtons.Controls.Add(Me.btnToTank)
        Me.pnlPipeButtons.Controls.Add(Me.btnCopyPipeProfile)
        Me.pnlPipeButtons.Controls.Add(Me.btnDeletePipe)
        Me.pnlPipeButtons.Controls.Add(Me.btnPipeSave)
        Me.pnlPipeButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlPipeButtons.Location = New System.Drawing.Point(0, 137)
        Me.pnlPipeButtons.Name = "pnlPipeButtons"
        Me.pnlPipeButtons.Size = New System.Drawing.Size(856, 40)
        Me.pnlPipeButtons.TabIndex = 1
        '
        'btnPipeCancel
        '
        Me.btnPipeCancel.Enabled = False
        Me.btnPipeCancel.Location = New System.Drawing.Point(171, 8)
        Me.btnPipeCancel.Name = "btnPipeCancel"
        Me.btnPipeCancel.Size = New System.Drawing.Size(75, 26)
        Me.btnPipeCancel.TabIndex = 1
        Me.btnPipeCancel.Text = "Cancel"
        '
        'btnPipeComments
        '
        Me.btnPipeComments.Location = New System.Drawing.Point(568, 8)
        Me.btnPipeComments.Name = "btnPipeComments"
        Me.btnPipeComments.Size = New System.Drawing.Size(88, 26)
        Me.btnPipeComments.TabIndex = 5
        Me.btnPipeComments.Text = "Comments"
        '
        'btnToTank
        '
        Me.btnToTank.Location = New System.Drawing.Point(472, 8)
        Me.btnToTank.Name = "btnToTank"
        Me.btnToTank.Size = New System.Drawing.Size(88, 26)
        Me.btnToTank.TabIndex = 4
        Me.btnToTank.Text = "Go To Tank"
        '
        'btnCopyPipeProfile
        '
        Me.btnCopyPipeProfile.Location = New System.Drawing.Point(250, 8)
        Me.btnCopyPipeProfile.Name = "btnCopyPipeProfile"
        Me.btnCopyPipeProfile.Size = New System.Drawing.Size(120, 26)
        Me.btnCopyPipeProfile.TabIndex = 2
        Me.btnCopyPipeProfile.Text = "Copy Pipe Profile"
        '
        'btnDeletePipe
        '
        Me.btnDeletePipe.Location = New System.Drawing.Point(376, 8)
        Me.btnDeletePipe.Name = "btnDeletePipe"
        Me.btnDeletePipe.Size = New System.Drawing.Size(88, 26)
        Me.btnDeletePipe.TabIndex = 3
        Me.btnDeletePipe.Text = "Delete Pipe"
        '
        'btnPipeSave
        '
        Me.btnPipeSave.Enabled = False
        Me.btnPipeSave.Location = New System.Drawing.Point(88, 8)
        Me.btnPipeSave.Name = "btnPipeSave"
        Me.btnPipeSave.Size = New System.Drawing.Size(80, 26)
        Me.btnPipeSave.TabIndex = 0
        Me.btnPipeSave.Text = "Save Pipe"
        '
        'pnlPipeDetailHeader
        '
        Me.pnlPipeDetailHeader.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblPipeTankIDValue)
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblPipeIDValue)
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblPipeIndex)
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblPipeCompartmentIndex)
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblPipeCompartment)
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblPipeID)
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblPipeTankID)
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblTankCountVal2)
        Me.pnlPipeDetailHeader.Controls.Add(Me.lblTankIDValue2)
        Me.pnlPipeDetailHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeDetailHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlPipeDetailHeader.Name = "pnlPipeDetailHeader"
        Me.pnlPipeDetailHeader.Size = New System.Drawing.Size(856, 24)
        Me.pnlPipeDetailHeader.TabIndex = 19
        '
        'lblPipeTankIDValue
        '
        Me.lblPipeTankIDValue.AutoSize = True
        Me.lblPipeTankIDValue.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblPipeTankIDValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblPipeTankIDValue.Location = New System.Drawing.Point(688, 4)
        Me.lblPipeTankIDValue.Name = "lblPipeTankIDValue"
        Me.lblPipeTankIDValue.Size = New System.Drawing.Size(0, 17)
        Me.lblPipeTankIDValue.TabIndex = 154
        Me.lblPipeTankIDValue.Visible = False
        '
        'lblPipeIDValue
        '
        Me.lblPipeIDValue.AutoSize = True
        Me.lblPipeIDValue.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblPipeIDValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblPipeIDValue.Location = New System.Drawing.Point(640, 4)
        Me.lblPipeIDValue.Name = "lblPipeIDValue"
        Me.lblPipeIDValue.Size = New System.Drawing.Size(0, 17)
        Me.lblPipeIDValue.TabIndex = 153
        Me.lblPipeIDValue.Visible = False
        '
        'lblPipeIndex
        '
        Me.lblPipeIndex.AutoSize = True
        Me.lblPipeIndex.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblPipeIndex.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblPipeIndex.Location = New System.Drawing.Point(528, 4)
        Me.lblPipeIndex.Name = "lblPipeIndex"
        Me.lblPipeIndex.Size = New System.Drawing.Size(11, 17)
        Me.lblPipeIndex.TabIndex = 151
        Me.lblPipeIndex.Text = "0"
        '
        'lblPipeCompartmentIndex
        '
        Me.lblPipeCompartmentIndex.AutoSize = True
        Me.lblPipeCompartmentIndex.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblPipeCompartmentIndex.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblPipeCompartmentIndex.Location = New System.Drawing.Point(404, 4)
        Me.lblPipeCompartmentIndex.Name = "lblPipeCompartmentIndex"
        Me.lblPipeCompartmentIndex.Size = New System.Drawing.Size(11, 17)
        Me.lblPipeCompartmentIndex.TabIndex = 150
        Me.lblPipeCompartmentIndex.Text = "0"
        '
        'lblPipeCompartment
        '
        Me.lblPipeCompartment.AutoSize = True
        Me.lblPipeCompartment.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblPipeCompartment.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeCompartment.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblPipeCompartment.Location = New System.Drawing.Point(304, 4)
        Me.lblPipeCompartment.Name = "lblPipeCompartment"
        Me.lblPipeCompartment.TabIndex = 149
        Me.lblPipeCompartment.Text = "Compartment #: "
        '
        'lblPipeID
        '
        Me.lblPipeID.AutoSize = True
        Me.lblPipeID.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblPipeID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeID.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblPipeID.Location = New System.Drawing.Point(480, 4)
        Me.lblPipeID.Name = "lblPipeID"
        Me.lblPipeID.Size = New System.Drawing.Size(48, 17)
        Me.lblPipeID.TabIndex = 148
        Me.lblPipeID.Text = "Pipe #: "
        '
        'lblPipeTankID
        '
        Me.lblPipeTankID.AutoSize = True
        Me.lblPipeTankID.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblPipeTankID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeTankID.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblPipeTankID.Location = New System.Drawing.Point(152, 4)
        Me.lblPipeTankID.Name = "lblPipeTankID"
        Me.lblPipeTankID.Size = New System.Drawing.Size(51, 17)
        Me.lblPipeTankID.TabIndex = 147
        Me.lblPipeTankID.Text = "Tank #: "
        '
        'lblTankCountVal2
        '
        Me.lblTankCountVal2.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblTankCountVal2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankCountVal2.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblTankCountVal2.Location = New System.Drawing.Point(232, 4)
        Me.lblTankCountVal2.Name = "lblTankCountVal2"
        Me.lblTankCountVal2.Size = New System.Drawing.Size(39, 17)
        Me.lblTankCountVal2.TabIndex = 208
        Me.lblTankCountVal2.Text = "of ???"
        '
        'lblTankIDValue2
        '
        Me.lblTankIDValue2.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblTankIDValue2.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblTankIDValue2.Location = New System.Drawing.Point(208, 4)
        Me.lblTankIDValue2.Name = "lblTankIDValue2"
        Me.lblTankIDValue2.Size = New System.Drawing.Size(18, 17)
        Me.lblTankIDValue2.TabIndex = 152
        Me.lblTankIDValue2.Text = "00"
        Me.lblTankIDValue2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'tbCntrlTank
        '
        Me.tbCntrlTank.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.tbCntrlTank.Controls.Add(Me.tbPageTankDetail)
        Me.tbCntrlTank.Location = New System.Drawing.Point(0, 424)
        Me.tbCntrlTank.Multiline = True
        Me.tbCntrlTank.Name = "tbCntrlTank"
        Me.tbCntrlTank.Padding = New System.Drawing.Point(0, 0)
        Me.tbCntrlTank.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tbCntrlTank.SelectedIndex = 0
        Me.tbCntrlTank.Size = New System.Drawing.Size(868, 464)
        Me.tbCntrlTank.TabIndex = 2
        Me.tbCntrlTank.Visible = False
        '
        'tbPageTankDetail
        '
        Me.tbPageTankDetail.Controls.Add(Me.pnlTankDetail)
        Me.tbPageTankDetail.Controls.Add(Me.pnlTankDetailHeader)
        Me.tbPageTankDetail.Controls.Add(Me.pnlTankButtons)
        Me.tbPageTankDetail.Location = New System.Drawing.Point(4, 27)
        Me.tbPageTankDetail.Name = "tbPageTankDetail"
        Me.tbPageTankDetail.Size = New System.Drawing.Size(860, 433)
        Me.tbPageTankDetail.TabIndex = 0
        Me.tbPageTankDetail.Text = "Tank Detail"
        '
        'pnlTankDetail
        '
        Me.pnlTankDetail.AutoScroll = True
        Me.pnlTankDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlTankDetail.Controls.Add(Me.pnlTankClosure)
        Me.pnlTankDetail.Controls.Add(Me.pnllblTankClosure)
        Me.pnlTankDetail.Controls.Add(Me.pnlTankInstallerOath)
        Me.pnlTankDetail.Controls.Add(Me.pnlInstallerOath)
        Me.pnlTankDetail.Controls.Add(Me.pnlTankRelease)
        Me.pnlTankDetail.Controls.Add(Me.pnlReleaseDetection)
        Me.pnlTankDetail.Controls.Add(Me.pnlTankMaterial)
        Me.pnlTankDetail.Controls.Add(Me.pnlTankMaterialHead)
        Me.pnlTankDetail.Controls.Add(Me.pnlTankTotalCapacity)
        Me.pnlTankDetail.Controls.Add(Me.Panel8)
        Me.pnlTankDetail.Controls.Add(Me.pnlTankInstallation)
        Me.pnlTankDetail.Controls.Add(Me.pnllblDateofInstallation)
        Me.pnlTankDetail.Controls.Add(Me.pnlTankDescriptionTop)
        Me.pnlTankDetail.Controls.Add(Me.pnlTankDescHead)
        Me.pnlTankDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlTankDetail.DockPadding.All = 2
        Me.pnlTankDetail.Location = New System.Drawing.Point(0, 24)
        Me.pnlTankDetail.Name = "pnlTankDetail"
        Me.pnlTankDetail.Size = New System.Drawing.Size(860, 369)
        Me.pnlTankDetail.TabIndex = 1
        '
        'pnlTankClosure
        '
        Me.pnlTankClosure.Controls.Add(Me.cmbTankInertFill)
        Me.pnlTankClosure.Controls.Add(Me.cmbTankClosureType)
        Me.pnlTankClosure.Controls.Add(Me.lblTankInertFillValue)
        Me.pnlTankClosure.Controls.Add(Me.lblTankInertFill)
        Me.pnlTankClosure.Controls.Add(Me.lblTankClosureStatusValue)
        Me.pnlTankClosure.Controls.Add(Me.lblTankClosureStatus)
        Me.pnlTankClosure.Controls.Add(Me.lblDateTankClosureRecvValue)
        Me.pnlTankClosure.Controls.Add(Me.lblDateClosureRecvd)
        Me.pnlTankClosure.Controls.Add(Me.lblDateClosed)
        Me.pnlTankClosure.Controls.Add(Me.dtPickLastUsed)
        Me.pnlTankClosure.Controls.Add(Me.lblDtLastUsed)
        Me.pnlTankClosure.Controls.Add(Me.lblClosuredate)
        Me.pnlTankClosure.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankClosure.Location = New System.Drawing.Point(2, 912)
        Me.pnlTankClosure.Name = "pnlTankClosure"
        Me.pnlTankClosure.Size = New System.Drawing.Size(852, 72)
        Me.pnlTankClosure.TabIndex = 6
        '
        'cmbTankInertFill
        '
        Me.cmbTankInertFill.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankInertFill.DropDownWidth = 300
        Me.cmbTankInertFill.Location = New System.Drawing.Point(520, 33)
        Me.cmbTankInertFill.Name = "cmbTankInertFill"
        Me.cmbTankInertFill.Size = New System.Drawing.Size(208, 23)
        Me.cmbTankInertFill.TabIndex = 2
        '
        'cmbTankClosureType
        '
        Me.cmbTankClosureType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankClosureType.DropDownWidth = 300
        Me.cmbTankClosureType.Location = New System.Drawing.Point(520, 8)
        Me.cmbTankClosureType.Name = "cmbTankClosureType"
        Me.cmbTankClosureType.Size = New System.Drawing.Size(208, 23)
        Me.cmbTankClosureType.TabIndex = 1
        '
        'lblTankInertFillValue
        '
        Me.lblTankInertFillValue.Location = New System.Drawing.Point(520, 34)
        Me.lblTankInertFillValue.Name = "lblTankInertFillValue"
        Me.lblTankInertFillValue.TabIndex = 182
        Me.lblTankInertFillValue.Text = "1"
        '
        'lblTankInertFill
        '
        Me.lblTankInertFill.Location = New System.Drawing.Point(456, 34)
        Me.lblTankInertFill.Name = "lblTankInertFill"
        Me.lblTankInertFill.Size = New System.Drawing.Size(53, 23)
        Me.lblTankInertFill.TabIndex = 181
        Me.lblTankInertFill.Text = "InertFill:"
        '
        'lblTankClosureStatusValue
        '
        Me.lblTankClosureStatusValue.Location = New System.Drawing.Point(520, 9)
        Me.lblTankClosureStatusValue.Name = "lblTankClosureStatusValue"
        Me.lblTankClosureStatusValue.TabIndex = 180
        Me.lblTankClosureStatusValue.Text = "0"
        '
        'lblTankClosureStatus
        '
        Me.lblTankClosureStatus.Location = New System.Drawing.Point(425, 9)
        Me.lblTankClosureStatus.Name = "lblTankClosureStatus"
        Me.lblTankClosureStatus.Size = New System.Drawing.Size(81, 21)
        Me.lblTankClosureStatus.TabIndex = 179
        Me.lblTankClosureStatus.Text = "Closure Type:"
        '
        'lblDateTankClosureRecvValue
        '
        Me.lblDateTankClosureRecvValue.Location = New System.Drawing.Point(136, 32)
        Me.lblDateTankClosureRecvValue.Name = "lblDateTankClosureRecvValue"
        Me.lblDateTankClosureRecvValue.TabIndex = 178
        '
        'lblDateClosureRecvd
        '
        Me.lblDateClosureRecvd.Location = New System.Drawing.Point(8, 32)
        Me.lblDateClosureRecvd.Name = "lblDateClosureRecvd"
        Me.lblDateClosureRecvd.Size = New System.Drawing.Size(111, 22)
        Me.lblDateClosureRecvd.TabIndex = 177
        Me.lblDateClosureRecvd.Text = "Date Closure Rcvd:"
        '
        'lblDateClosed
        '
        Me.lblDateClosed.Location = New System.Drawing.Point(42, 56)
        Me.lblDateClosed.Name = "lblDateClosed"
        Me.lblDateClosed.Size = New System.Drawing.Size(80, 23)
        Me.lblDateClosed.TabIndex = 174
        Me.lblDateClosed.Text = "Date Closed:"
        '
        'dtPickLastUsed
        '
        Me.dtPickLastUsed.Checked = False
        Me.dtPickLastUsed.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickLastUsed.Location = New System.Drawing.Point(136, 8)
        Me.dtPickLastUsed.Name = "dtPickLastUsed"
        Me.dtPickLastUsed.ShowCheckBox = True
        Me.dtPickLastUsed.Size = New System.Drawing.Size(104, 21)
        Me.dtPickLastUsed.TabIndex = 0
        '
        'lblDtLastUsed
        '
        Me.lblDtLastUsed.AutoSize = True
        Me.lblDtLastUsed.Location = New System.Drawing.Point(27, 8)
        Me.lblDtLastUsed.Name = "lblDtLastUsed"
        Me.lblDtLastUsed.Size = New System.Drawing.Size(93, 17)
        Me.lblDtLastUsed.TabIndex = 173
        Me.lblDtLastUsed.Text = "Date Last Used:"
        '
        'lblClosuredate
        '
        Me.lblClosuredate.BackColor = System.Drawing.Color.Red
        Me.lblClosuredate.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblClosuredate.Location = New System.Drawing.Point(136, 56)
        Me.lblClosuredate.Name = "lblClosuredate"
        Me.lblClosuredate.Size = New System.Drawing.Size(104, 17)
        Me.lblClosuredate.TabIndex = 175
        '
        'pnllblTankClosure
        '
        Me.pnllblTankClosure.Controls.Add(Me.lblTankClosureCaption)
        Me.pnllblTankClosure.Controls.Add(Me.lblTankClosure)
        Me.pnllblTankClosure.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnllblTankClosure.Location = New System.Drawing.Point(2, 888)
        Me.pnllblTankClosure.Name = "pnllblTankClosure"
        Me.pnllblTankClosure.Size = New System.Drawing.Size(852, 24)
        Me.pnllblTankClosure.TabIndex = 191
        '
        'lblTankClosureCaption
        '
        Me.lblTankClosureCaption.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTankClosureCaption.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTankClosureCaption.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTankClosureCaption.Location = New System.Drawing.Point(16, 0)
        Me.lblTankClosureCaption.Name = "lblTankClosureCaption"
        Me.lblTankClosureCaption.Size = New System.Drawing.Size(836, 24)
        Me.lblTankClosureCaption.TabIndex = 3
        Me.lblTankClosureCaption.Text = "Closing of Tank"
        Me.lblTankClosureCaption.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTankClosure
        '
        Me.lblTankClosure.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTankClosure.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTankClosure.Location = New System.Drawing.Point(0, 0)
        Me.lblTankClosure.Name = "lblTankClosure"
        Me.lblTankClosure.Size = New System.Drawing.Size(16, 24)
        Me.lblTankClosure.TabIndex = 2
        Me.lblTankClosure.Text = "-"
        Me.lblTankClosure.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlTankInstallerOath
        '
        Me.pnlTankInstallerOath.Controls.Add(Me.txtLicensee)
        Me.pnlTankInstallerOath.Controls.Add(Me.lblLicenseeSearch)
        Me.pnlTankInstallerOath.Controls.Add(Me.txtTankCompany)
        Me.pnlTankInstallerOath.Controls.Add(Me.lblLicenseeName)
        Me.pnlTankInstallerOath.Controls.Add(Me.lblTankInstallerCompany)
        Me.pnlTankInstallerOath.Controls.Add(Me.dtPickTankInstallerSigned)
        Me.pnlTankInstallerOath.Controls.Add(Me.lblTankInstallerDtSigned)
        Me.pnlTankInstallerOath.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankInstallerOath.Location = New System.Drawing.Point(2, 792)
        Me.pnlTankInstallerOath.Name = "pnlTankInstallerOath"
        Me.pnlTankInstallerOath.Size = New System.Drawing.Size(852, 96)
        Me.pnlTankInstallerOath.TabIndex = 5
        '
        'txtLicensee
        '
        Me.txtLicensee.Location = New System.Drawing.Point(200, 32)
        Me.txtLicensee.Name = "txtLicensee"
        Me.txtLicensee.Size = New System.Drawing.Size(200, 21)
        Me.txtLicensee.TabIndex = 221
        Me.txtLicensee.Text = ""
        '
        'lblLicenseeSearch
        '
        Me.lblLicenseeSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLicenseeSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLicenseeSearch.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblLicenseeSearch.Location = New System.Drawing.Point(408, 32)
        Me.lblLicenseeSearch.Name = "lblLicenseeSearch"
        Me.lblLicenseeSearch.Size = New System.Drawing.Size(16, 22)
        Me.lblLicenseeSearch.TabIndex = 220
        Me.lblLicenseeSearch.Text = "?"
        Me.lblLicenseeSearch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtTankCompany
        '
        Me.txtTankCompany.Location = New System.Drawing.Point(200, 56)
        Me.txtTankCompany.Name = "txtTankCompany"
        Me.txtTankCompany.Size = New System.Drawing.Size(200, 21)
        Me.txtTankCompany.TabIndex = 2
        Me.txtTankCompany.Text = ""
        '
        'lblLicenseeName
        '
        Me.lblLicenseeName.Location = New System.Drawing.Point(128, 32)
        Me.lblLicenseeName.Name = "lblLicenseeName"
        Me.lblLicenseeName.Size = New System.Drawing.Size(64, 17)
        Me.lblLicenseeName.TabIndex = 190
        Me.lblLicenseeName.Text = "Licensee:"
        Me.lblLicenseeName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTankInstallerCompany
        '
        Me.lblTankInstallerCompany.Location = New System.Drawing.Point(128, 56)
        Me.lblTankInstallerCompany.Name = "lblTankInstallerCompany"
        Me.lblTankInstallerCompany.Size = New System.Drawing.Size(64, 17)
        Me.lblTankInstallerCompany.TabIndex = 191
        Me.lblTankInstallerCompany.Text = "Company:"
        Me.lblTankInstallerCompany.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtPickTankInstallerSigned
        '
        Me.dtPickTankInstallerSigned.Checked = False
        Me.dtPickTankInstallerSigned.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickTankInstallerSigned.Location = New System.Drawing.Point(200, 4)
        Me.dtPickTankInstallerSigned.Name = "dtPickTankInstallerSigned"
        Me.dtPickTankInstallerSigned.ShowCheckBox = True
        Me.dtPickTankInstallerSigned.Size = New System.Drawing.Size(103, 21)
        Me.dtPickTankInstallerSigned.TabIndex = 0
        '
        'lblTankInstallerDtSigned
        '
        Me.lblTankInstallerDtSigned.AutoSize = True
        Me.lblTankInstallerDtSigned.Location = New System.Drawing.Point(120, 8)
        Me.lblTankInstallerDtSigned.Name = "lblTankInstallerDtSigned"
        Me.lblTankInstallerDtSigned.Size = New System.Drawing.Size(76, 17)
        Me.lblTankInstallerDtSigned.TabIndex = 127
        Me.lblTankInstallerDtSigned.Text = "Date Signed:"
        '
        'pnlInstallerOath
        '
        Me.pnlInstallerOath.Controls.Add(Me.lblTankInstallerOath)
        Me.pnlInstallerOath.Controls.Add(Me.lblTankInstallerOathDisplay)
        Me.pnlInstallerOath.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlInstallerOath.Location = New System.Drawing.Point(2, 768)
        Me.pnlInstallerOath.Name = "pnlInstallerOath"
        Me.pnlInstallerOath.Size = New System.Drawing.Size(852, 24)
        Me.pnlInstallerOath.TabIndex = 185
        '
        'lblTankInstallerOath
        '
        Me.lblTankInstallerOath.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTankInstallerOath.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTankInstallerOath.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTankInstallerOath.Location = New System.Drawing.Point(16, 0)
        Me.lblTankInstallerOath.Name = "lblTankInstallerOath"
        Me.lblTankInstallerOath.Size = New System.Drawing.Size(836, 24)
        Me.lblTankInstallerOath.TabIndex = 1
        Me.lblTankInstallerOath.Text = "Installer Oath"
        Me.lblTankInstallerOath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTankInstallerOathDisplay
        '
        Me.lblTankInstallerOathDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTankInstallerOathDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTankInstallerOathDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblTankInstallerOathDisplay.Name = "lblTankInstallerOathDisplay"
        Me.lblTankInstallerOathDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblTankInstallerOathDisplay.TabIndex = 0
        Me.lblTankInstallerOathDisplay.Text = "-"
        Me.lblTankInstallerOathDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlTankRelease
        '
        Me.pnlTankRelease.Controls.Add(Me.dtPickATGLastInspected)
        Me.pnlTankRelease.Controls.Add(Me.dtPickElectronicDeviceInspected)
        Me.pnlTankRelease.Controls.Add(Me.LblDateATGLastInspected)
        Me.pnlTankRelease.Controls.Add(Me.LblDateElectronicDeviceInspected)
        Me.pnlTankRelease.Controls.Add(Me.cmbTankReleaseDetection)
        Me.pnlTankRelease.Controls.Add(Me.dtPickTankTightnessTest)
        Me.pnlTankRelease.Controls.Add(Me.lblRelseDetection)
        Me.pnlTankRelease.Controls.Add(Me.lblLTankTightnessTstDt)
        Me.pnlTankRelease.Controls.Add(Me.chkTankDrpTubeInvControl)
        Me.pnlTankRelease.Controls.Add(Me.dtPickSecondaryContainmentLastInspected)
        Me.pnlTankRelease.Controls.Add(Me.lbLDateSecondaryContainmentLastInspected)
        Me.pnlTankRelease.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankRelease.Location = New System.Drawing.Point(2, 656)
        Me.pnlTankRelease.Name = "pnlTankRelease"
        Me.pnlTankRelease.Size = New System.Drawing.Size(852, 112)
        Me.pnlTankRelease.TabIndex = 4
        '
        'dtPickATGLastInspected
        '
        Me.dtPickATGLastInspected.Checked = False
        Me.dtPickATGLastInspected.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickATGLastInspected.Location = New System.Drawing.Point(528, 57)
        Me.dtPickATGLastInspected.Name = "dtPickATGLastInspected"
        Me.dtPickATGLastInspected.ShowCheckBox = True
        Me.dtPickATGLastInspected.Size = New System.Drawing.Size(104, 21)
        Me.dtPickATGLastInspected.TabIndex = 165
        '
        'dtPickElectronicDeviceInspected
        '
        Me.dtPickElectronicDeviceInspected.Checked = False
        Me.dtPickElectronicDeviceInspected.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickElectronicDeviceInspected.Location = New System.Drawing.Point(528, 34)
        Me.dtPickElectronicDeviceInspected.Name = "dtPickElectronicDeviceInspected"
        Me.dtPickElectronicDeviceInspected.ShowCheckBox = True
        Me.dtPickElectronicDeviceInspected.Size = New System.Drawing.Size(104, 21)
        Me.dtPickElectronicDeviceInspected.TabIndex = 164
        '
        'LblDateATGLastInspected
        '
        Me.LblDateATGLastInspected.AutoSize = True
        Me.LblDateATGLastInspected.Location = New System.Drawing.Point(328, 57)
        Me.LblDateATGLastInspected.Name = "LblDateATGLastInspected"
        Me.LblDateATGLastInspected.Size = New System.Drawing.Size(143, 17)
        Me.LblDateATGLastInspected.TabIndex = 158
        Me.LblDateATGLastInspected.Text = "Date ATG Last Inspected"
        Me.LblDateATGLastInspected.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LblDateElectronicDeviceInspected
        '
        Me.LblDateElectronicDeviceInspected.AutoSize = True
        Me.LblDateElectronicDeviceInspected.Location = New System.Drawing.Point(328, 34)
        Me.LblDateElectronicDeviceInspected.Name = "LblDateElectronicDeviceInspected"
        Me.LblDateElectronicDeviceInspected.Size = New System.Drawing.Size(186, 17)
        Me.LblDateElectronicDeviceInspected.TabIndex = 156
        Me.LblDateElectronicDeviceInspected.Text = "Date Electronic Device Inspected"
        Me.LblDateElectronicDeviceInspected.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbTankReleaseDetection
        '
        Me.cmbTankReleaseDetection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankReleaseDetection.DropDownWidth = 300
        Me.cmbTankReleaseDetection.Location = New System.Drawing.Point(192, 5)
        Me.cmbTankReleaseDetection.Name = "cmbTankReleaseDetection"
        Me.cmbTankReleaseDetection.Size = New System.Drawing.Size(328, 23)
        Me.cmbTankReleaseDetection.TabIndex = 0
        '
        'dtPickTankTightnessTest
        '
        Me.dtPickTankTightnessTest.Checked = False
        Me.dtPickTankTightnessTest.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickTankTightnessTest.Location = New System.Drawing.Point(192, 32)
        Me.dtPickTankTightnessTest.Name = "dtPickTankTightnessTest"
        Me.dtPickTankTightnessTest.ShowCheckBox = True
        Me.dtPickTankTightnessTest.Size = New System.Drawing.Size(104, 21)
        Me.dtPickTankTightnessTest.TabIndex = 1
        '
        'lblRelseDetection
        '
        Me.lblRelseDetection.AutoSize = True
        Me.lblRelseDetection.Location = New System.Drawing.Point(80, 8)
        Me.lblRelseDetection.Name = "lblRelseDetection"
        Me.lblRelseDetection.Size = New System.Drawing.Size(109, 17)
        Me.lblRelseDetection.TabIndex = 133
        Me.lblRelseDetection.Text = "Release Detection:"
        '
        'lblLTankTightnessTstDt
        '
        Me.lblLTankTightnessTstDt.AutoSize = True
        Me.lblLTankTightnessTstDt.Location = New System.Drawing.Point(16, 32)
        Me.lblLTankTightnessTstDt.Name = "lblLTankTightnessTstDt"
        Me.lblLTankTightnessTstDt.Size = New System.Drawing.Size(174, 17)
        Me.lblLTankTightnessTstDt.TabIndex = 134
        Me.lblLTankTightnessTstDt.Text = "Last Tank Tightness Test date:"
        '
        'chkTankDrpTubeInvControl
        '
        Me.chkTankDrpTubeInvControl.Location = New System.Drawing.Point(22, 56)
        Me.chkTankDrpTubeInvControl.Name = "chkTankDrpTubeInvControl"
        Me.chkTankDrpTubeInvControl.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkTankDrpTubeInvControl.Size = New System.Drawing.Size(184, 24)
        Me.chkTankDrpTubeInvControl.TabIndex = 2
        Me.chkTankDrpTubeInvControl.Text = "Drop Tube Inventory Control"
        Me.chkTankDrpTubeInvControl.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dtPickSecondaryContainmentLastInspected
        '
        Me.dtPickSecondaryContainmentLastInspected.Checked = False
        Me.dtPickSecondaryContainmentLastInspected.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickSecondaryContainmentLastInspected.Location = New System.Drawing.Point(528, 80)
        Me.dtPickSecondaryContainmentLastInspected.Name = "dtPickSecondaryContainmentLastInspected"
        Me.dtPickSecondaryContainmentLastInspected.ShowCheckBox = True
        Me.dtPickSecondaryContainmentLastInspected.Size = New System.Drawing.Size(104, 21)
        Me.dtPickSecondaryContainmentLastInspected.TabIndex = 163
        '
        'lbLDateSecondaryContainmentLastInspected
        '
        Me.lbLDateSecondaryContainmentLastInspected.AutoSize = True
        Me.lbLDateSecondaryContainmentLastInspected.Location = New System.Drawing.Point(272, 80)
        Me.lbLDateSecondaryContainmentLastInspected.Name = "lbLDateSecondaryContainmentLastInspected"
        Me.lbLDateSecondaryContainmentLastInspected.Size = New System.Drawing.Size(253, 17)
        Me.lbLDateSecondaryContainmentLastInspected.TabIndex = 160
        Me.lbLDateSecondaryContainmentLastInspected.Text = "Date Secondary Containment Last Inspected:"
        Me.lbLDateSecondaryContainmentLastInspected.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlReleaseDetection
        '
        Me.pnlReleaseDetection.Controls.Add(Me.lblTankReleaseHead)
        Me.pnlReleaseDetection.Controls.Add(Me.lblTankReleaseDisplay)
        Me.pnlReleaseDetection.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlReleaseDetection.Location = New System.Drawing.Point(2, 632)
        Me.pnlReleaseDetection.Name = "pnlReleaseDetection"
        Me.pnlReleaseDetection.Size = New System.Drawing.Size(852, 24)
        Me.pnlReleaseDetection.TabIndex = 183
        '
        'lblTankReleaseHead
        '
        Me.lblTankReleaseHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTankReleaseHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTankReleaseHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTankReleaseHead.Location = New System.Drawing.Point(16, 0)
        Me.lblTankReleaseHead.Name = "lblTankReleaseHead"
        Me.lblTankReleaseHead.Size = New System.Drawing.Size(836, 24)
        Me.lblTankReleaseHead.TabIndex = 1
        Me.lblTankReleaseHead.Text = "Release Detection"
        Me.lblTankReleaseHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTankReleaseDisplay
        '
        Me.lblTankReleaseDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTankReleaseDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTankReleaseDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblTankReleaseDisplay.Name = "lblTankReleaseDisplay"
        Me.lblTankReleaseDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblTankReleaseDisplay.TabIndex = 0
        Me.lblTankReleaseDisplay.Text = "-"
        Me.lblTankReleaseDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlTankMaterial
        '
        Me.pnlTankMaterial.Controls.Add(Me.dtPickOverfillPreventionInstalled)
        Me.pnlTankMaterial.Controls.Add(Me.LblDateOverfillPreventionInstalled)
        Me.pnlTankMaterial.Controls.Add(Me.dtPickOverfillPreventionLastInspected)
        Me.pnlTankMaterial.Controls.Add(Me.dtPickSpillPreventionLastTested)
        Me.pnlTankMaterial.Controls.Add(Me.LblDateOverfillPreventionLastInspected)
        Me.pnlTankMaterial.Controls.Add(Me.LblDateSpillPreventionLastTested)
        Me.pnlTankMaterial.Controls.Add(Me.dtPickSpillPreventionInstalled)
        Me.pnlTankMaterial.Controls.Add(Me.LblDateSpillPreventionInstalled)
        Me.pnlTankMaterial.Controls.Add(Me.chkDeliveriesLimited)
        Me.pnlTankMaterial.Controls.Add(Me.dtPickCPLastTested)
        Me.pnlTankMaterial.Controls.Add(Me.dtPickCPInstalled)
        Me.pnlTankMaterial.Controls.Add(Me.cmbTankOverfillProtectionType)
        Me.pnlTankMaterial.Controls.Add(Me.dtPickInteriorLiningInstalled)
        Me.pnlTankMaterial.Controls.Add(Me.dtPickLastInteriorLinningInspection)
        Me.pnlTankMaterial.Controls.Add(Me.lblTankMaterial)
        Me.pnlTankMaterial.Controls.Add(Me.lblTankOption)
        Me.pnlTankMaterial.Controls.Add(Me.cmbTankMaterial)
        Me.pnlTankMaterial.Controls.Add(Me.cmbTankOptions)
        Me.pnlTankMaterial.Controls.Add(Me.lblTankCPType)
        Me.pnlTankMaterial.Controls.Add(Me.cmbTankCPType)
        Me.pnlTankMaterial.Controls.Add(Me.lblDateCPInstalled)
        Me.pnlTankMaterial.Controls.Add(Me.lblDtTankLstTested)
        Me.pnlTankMaterial.Controls.Add(Me.chkOverFilledProtected)
        Me.pnlTankMaterial.Controls.Add(Me.chkBxSpillProtected)
        Me.pnlTankMaterial.Controls.Add(Me.chkBxTightfillAdapters)
        Me.pnlTankMaterial.Controls.Add(Me.lblDateLnInteriorInstalled)
        Me.pnlTankMaterial.Controls.Add(Me.lblDtLnInteriorLstInspect)
        Me.pnlTankMaterial.Controls.Add(Me.lblTankOverfillProtectionType)
        Me.pnlTankMaterial.Controls.Add(Me.chkEmergencyPower)
        Me.pnlTankMaterial.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankMaterial.DockPadding.All = 100
        Me.pnlTankMaterial.Location = New System.Drawing.Point(2, 440)
        Me.pnlTankMaterial.Name = "pnlTankMaterial"
        Me.pnlTankMaterial.Size = New System.Drawing.Size(852, 192)
        Me.pnlTankMaterial.TabIndex = 3
        '
        'dtPickOverfillPreventionInstalled
        '
        Me.dtPickOverfillPreventionInstalled.Checked = False
        Me.dtPickOverfillPreventionInstalled.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickOverfillPreventionInstalled.Location = New System.Drawing.Point(624, 133)
        Me.dtPickOverfillPreventionInstalled.Name = "dtPickOverfillPreventionInstalled"
        Me.dtPickOverfillPreventionInstalled.ShowCheckBox = True
        Me.dtPickOverfillPreventionInstalled.Size = New System.Drawing.Size(104, 21)
        Me.dtPickOverfillPreventionInstalled.TabIndex = 165
        '
        'LblDateOverfillPreventionInstalled
        '
        Me.LblDateOverfillPreventionInstalled.AutoSize = True
        Me.LblDateOverfillPreventionInstalled.Location = New System.Drawing.Point(424, 133)
        Me.LblDateOverfillPreventionInstalled.Name = "LblDateOverfillPreventionInstalled"
        Me.LblDateOverfillPreventionInstalled.Size = New System.Drawing.Size(188, 17)
        Me.LblDateOverfillPreventionInstalled.TabIndex = 164
        Me.LblDateOverfillPreventionInstalled.Text = "Date Overfill Prevention Installed:"
        Me.LblDateOverfillPreventionInstalled.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtPickOverfillPreventionLastInspected
        '
        Me.dtPickOverfillPreventionLastInspected.Checked = False
        Me.dtPickOverfillPreventionLastInspected.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickOverfillPreventionLastInspected.Location = New System.Drawing.Point(624, 160)
        Me.dtPickOverfillPreventionLastInspected.Name = "dtPickOverfillPreventionLastInspected"
        Me.dtPickOverfillPreventionLastInspected.ShowCheckBox = True
        Me.dtPickOverfillPreventionLastInspected.Size = New System.Drawing.Size(104, 21)
        Me.dtPickOverfillPreventionLastInspected.TabIndex = 162
        '
        'dtPickSpillPreventionLastTested
        '
        Me.dtPickSpillPreventionLastTested.Checked = False
        Me.dtPickSpillPreventionLastTested.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickSpillPreventionLastTested.Location = New System.Drawing.Point(232, 160)
        Me.dtPickSpillPreventionLastTested.Name = "dtPickSpillPreventionLastTested"
        Me.dtPickSpillPreventionLastTested.ShowCheckBox = True
        Me.dtPickSpillPreventionLastTested.Size = New System.Drawing.Size(104, 21)
        Me.dtPickSpillPreventionLastTested.TabIndex = 161
        '
        'LblDateOverfillPreventionLastInspected
        '
        Me.LblDateOverfillPreventionLastInspected.AutoSize = True
        Me.LblDateOverfillPreventionLastInspected.Location = New System.Drawing.Point(392, 160)
        Me.LblDateOverfillPreventionLastInspected.Name = "LblDateOverfillPreventionLastInspected"
        Me.LblDateOverfillPreventionLastInspected.Size = New System.Drawing.Size(222, 17)
        Me.LblDateOverfillPreventionLastInspected.TabIndex = 158
        Me.LblDateOverfillPreventionLastInspected.Text = "Date Overfill Prevention Last Inspected:"
        Me.LblDateOverfillPreventionLastInspected.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LblDateSpillPreventionLastTested
        '
        Me.LblDateSpillPreventionLastTested.AutoSize = True
        Me.LblDateSpillPreventionLastTested.Location = New System.Drawing.Point(40, 160)
        Me.LblDateSpillPreventionLastTested.Name = "LblDateSpillPreventionLastTested"
        Me.LblDateSpillPreventionLastTested.Size = New System.Drawing.Size(191, 17)
        Me.LblDateSpillPreventionLastTested.TabIndex = 156
        Me.LblDateSpillPreventionLastTested.Text = "Date Spill Prevention Last Tested:"
        Me.LblDateSpillPreventionLastTested.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtPickSpillPreventionInstalled
        '
        Me.dtPickSpillPreventionInstalled.Checked = False
        Me.dtPickSpillPreventionInstalled.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickSpillPreventionInstalled.Location = New System.Drawing.Point(232, 136)
        Me.dtPickSpillPreventionInstalled.Name = "dtPickSpillPreventionInstalled"
        Me.dtPickSpillPreventionInstalled.ShowCheckBox = True
        Me.dtPickSpillPreventionInstalled.Size = New System.Drawing.Size(104, 21)
        Me.dtPickSpillPreventionInstalled.TabIndex = 153
        '
        'LblDateSpillPreventionInstalled
        '
        Me.LblDateSpillPreventionInstalled.AutoSize = True
        Me.LblDateSpillPreventionInstalled.Location = New System.Drawing.Point(56, 136)
        Me.LblDateSpillPreventionInstalled.Name = "LblDateSpillPreventionInstalled"
        Me.LblDateSpillPreventionInstalled.Size = New System.Drawing.Size(173, 17)
        Me.LblDateSpillPreventionInstalled.TabIndex = 154
        Me.LblDateSpillPreventionInstalled.Text = "Date Spill Prevention Installed:"
        Me.LblDateSpillPreventionInstalled.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkDeliveriesLimited
        '
        Me.chkDeliveriesLimited.Location = New System.Drawing.Point(477, 32)
        Me.chkDeliveriesLimited.Name = "chkDeliveriesLimited"
        Me.chkDeliveriesLimited.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkDeliveriesLimited.Size = New System.Drawing.Size(247, 24)
        Me.chkDeliveriesLimited.TabIndex = 7
        Me.chkDeliveriesLimited.Text = "If Deliveries limited to 25 gallons or less"
        Me.chkDeliveriesLimited.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtPickCPLastTested
        '
        Me.dtPickCPLastTested.Checked = False
        Me.dtPickCPLastTested.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickCPLastTested.Location = New System.Drawing.Point(232, 112)
        Me.dtPickCPLastTested.Name = "dtPickCPLastTested"
        Me.dtPickCPLastTested.ShowCheckBox = True
        Me.dtPickCPLastTested.Size = New System.Drawing.Size(104, 21)
        Me.dtPickCPLastTested.TabIndex = 4
        '
        'dtPickCPInstalled
        '
        Me.dtPickCPInstalled.Checked = False
        Me.dtPickCPInstalled.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickCPInstalled.Location = New System.Drawing.Point(232, 86)
        Me.dtPickCPInstalled.Name = "dtPickCPInstalled"
        Me.dtPickCPInstalled.ShowCheckBox = True
        Me.dtPickCPInstalled.Size = New System.Drawing.Size(104, 21)
        Me.dtPickCPInstalled.TabIndex = 3
        '
        'cmbTankOverfillProtectionType
        '
        Me.cmbTankOverfillProtectionType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankOverfillProtectionType.Location = New System.Drawing.Point(592, 104)
        Me.cmbTankOverfillProtectionType.Name = "cmbTankOverfillProtectionType"
        Me.cmbTankOverfillProtectionType.Size = New System.Drawing.Size(136, 23)
        Me.cmbTankOverfillProtectionType.TabIndex = 10
        '
        'dtPickInteriorLiningInstalled
        '
        Me.dtPickInteriorLiningInstalled.Checked = False
        Me.dtPickInteriorLiningInstalled.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickInteriorLiningInstalled.Location = New System.Drawing.Point(625, 57)
        Me.dtPickInteriorLiningInstalled.Name = "dtPickInteriorLiningInstalled"
        Me.dtPickInteriorLiningInstalled.ShowCheckBox = True
        Me.dtPickInteriorLiningInstalled.Size = New System.Drawing.Size(103, 21)
        Me.dtPickInteriorLiningInstalled.TabIndex = 8
        '
        'dtPickLastInteriorLinningInspection
        '
        Me.dtPickLastInteriorLinningInspection.Checked = False
        Me.dtPickLastInteriorLinningInspection.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickLastInteriorLinningInspection.Location = New System.Drawing.Point(625, 81)
        Me.dtPickLastInteriorLinningInspection.Name = "dtPickLastInteriorLinningInspection"
        Me.dtPickLastInteriorLinningInspection.ShowCheckBox = True
        Me.dtPickLastInteriorLinningInspection.Size = New System.Drawing.Size(103, 21)
        Me.dtPickLastInteriorLinningInspection.TabIndex = 9
        '
        'lblTankMaterial
        '
        Me.lblTankMaterial.AutoSize = True
        Me.lblTankMaterial.Location = New System.Drawing.Point(144, 8)
        Me.lblTankMaterial.Name = "lblTankMaterial"
        Me.lblTankMaterial.Size = New System.Drawing.Size(83, 17)
        Me.lblTankMaterial.TabIndex = 144
        Me.lblTankMaterial.Text = "Tank Material:"
        Me.lblTankMaterial.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTankOption
        '
        Me.lblTankOption.AutoSize = True
        Me.lblTankOption.Location = New System.Drawing.Point(88, 32)
        Me.lblTankOption.Name = "lblTankOption"
        Me.lblTankOption.Size = New System.Drawing.Size(138, 17)
        Me.lblTankOption.TabIndex = 145
        Me.lblTankOption.Text = "Secondary Tank Option:"
        Me.lblTankOption.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbTankMaterial
        '
        Me.cmbTankMaterial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankMaterial.DropDownWidth = 300
        Me.cmbTankMaterial.Items.AddRange(New Object() {"Asphalt Coated or Bare Steel", "Composite (Steel w/FRP)", "Composite (Steel w/Plastic)", "Epoxy Coated Steel", "Fiberglass Reinforced Plastic", "Unknown", "Other"})
        Me.cmbTankMaterial.Location = New System.Drawing.Point(232, 8)
        Me.cmbTankMaterial.Name = "cmbTankMaterial"
        Me.cmbTankMaterial.Size = New System.Drawing.Size(210, 23)
        Me.cmbTankMaterial.TabIndex = 0
        '
        'cmbTankOptions
        '
        Me.cmbTankOptions.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankOptions.DropDownWidth = 300
        Me.cmbTankOptions.Items.AddRange(New Object() {"Double-Walled", "Lined Interior", "None", "Cathodically Protected", "Double-walled/Cathodically Protected", "Cathodically Protected/Lined Interior"})
        Me.cmbTankOptions.Location = New System.Drawing.Point(232, 32)
        Me.cmbTankOptions.Name = "cmbTankOptions"
        Me.cmbTankOptions.Size = New System.Drawing.Size(210, 23)
        Me.cmbTankOptions.TabIndex = 1
        '
        'lblTankCPType
        '
        Me.lblTankCPType.AutoSize = True
        Me.lblTankCPType.Location = New System.Drawing.Point(136, 56)
        Me.lblTankCPType.Name = "lblTankCPType"
        Me.lblTankCPType.Size = New System.Drawing.Size(87, 17)
        Me.lblTankCPType.TabIndex = 146
        Me.lblTankCPType.Text = "Tank CP Type:"
        Me.lblTankCPType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbTankCPType
        '
        Me.cmbTankCPType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankCPType.DropDownWidth = 300
        Me.cmbTankCPType.Location = New System.Drawing.Point(232, 56)
        Me.cmbTankCPType.Name = "cmbTankCPType"
        Me.cmbTankCPType.Size = New System.Drawing.Size(210, 23)
        Me.cmbTankCPType.TabIndex = 2
        '
        'lblDateCPInstalled
        '
        Me.lblDateCPInstalled.AutoSize = True
        Me.lblDateCPInstalled.Location = New System.Drawing.Point(120, 86)
        Me.lblDateCPInstalled.Name = "lblDateCPInstalled"
        Me.lblDateCPInstalled.Size = New System.Drawing.Size(104, 17)
        Me.lblDateCPInstalled.TabIndex = 147
        Me.lblDateCPInstalled.Text = "Date CP Installed:"
        Me.lblDateCPInstalled.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDtTankLstTested
        '
        Me.lblDtTankLstTested.AutoSize = True
        Me.lblDtTankLstTested.Location = New System.Drawing.Point(72, 112)
        Me.lblDtTankLstTested.Name = "lblDtTankLstTested"
        Me.lblDtTankLstTested.Size = New System.Drawing.Size(153, 17)
        Me.lblDtTankLstTested.TabIndex = 149
        Me.lblDtTankLstTested.Text = "Date Tank CP Last Tested:"
        Me.lblDtTankLstTested.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkOverFilledProtected
        '
        Me.chkOverFilledProtected.Location = New System.Drawing.Point(736, 64)
        Me.chkOverFilledProtected.Name = "chkOverFilledProtected"
        Me.chkOverFilledProtected.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkOverFilledProtected.Size = New System.Drawing.Size(120, 24)
        Me.chkOverFilledProtected.TabIndex = 12
        Me.chkOverFilledProtected.Text = "Overfill Protected"
        Me.chkOverFilledProtected.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.chkOverFilledProtected.Visible = False
        '
        'chkBxSpillProtected
        '
        Me.chkBxSpillProtected.Location = New System.Drawing.Point(736, 16)
        Me.chkBxSpillProtected.Name = "chkBxSpillProtected"
        Me.chkBxSpillProtected.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkBxSpillProtected.TabIndex = 5
        Me.chkBxSpillProtected.Text = "Spill Protected"
        Me.chkBxSpillProtected.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.chkBxSpillProtected.Visible = False
        '
        'chkBxTightfillAdapters
        '
        Me.chkBxTightfillAdapters.Location = New System.Drawing.Point(736, 40)
        Me.chkBxTightfillAdapters.Name = "chkBxTightfillAdapters"
        Me.chkBxTightfillAdapters.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkBxTightfillAdapters.Size = New System.Drawing.Size(120, 24)
        Me.chkBxTightfillAdapters.TabIndex = 11
        Me.chkBxTightfillAdapters.Text = "Tightfill Adapters"
        Me.chkBxTightfillAdapters.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.chkBxTightfillAdapters.Visible = False
        '
        'lblDateLnInteriorInstalled
        '
        Me.lblDateLnInteriorInstalled.AutoSize = True
        Me.lblDateLnInteriorInstalled.Location = New System.Drawing.Point(462, 57)
        Me.lblDateLnInteriorInstalled.Name = "lblDateLnInteriorInstalled"
        Me.lblDateLnInteriorInstalled.Size = New System.Drawing.Size(155, 17)
        Me.lblDateLnInteriorInstalled.TabIndex = 148
        Me.lblDateLnInteriorInstalled.Text = "Interior Lining Installed  On:"
        Me.lblDateLnInteriorInstalled.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDtLnInteriorLstInspect
        '
        Me.lblDtLnInteriorLstInspect.AutoSize = True
        Me.lblDtLnInteriorLstInspect.Location = New System.Drawing.Point(408, 81)
        Me.lblDtLnInteriorLstInspect.Name = "lblDtLnInteriorLstInspect"
        Me.lblDtLnInteriorLstInspect.Size = New System.Drawing.Size(209, 17)
        Me.lblDtLnInteriorLstInspect.TabIndex = 150
        Me.lblDtLnInteriorLstInspect.Text = "Date of Last Interior Lining Inspection"
        '
        'lblTankOverfillProtectionType
        '
        Me.lblTankOverfillProtectionType.AutoSize = True
        Me.lblTankOverfillProtectionType.Location = New System.Drawing.Point(448, 104)
        Me.lblTankOverfillProtectionType.Name = "lblTankOverfillProtectionType"
        Me.lblTankOverfillProtectionType.Size = New System.Drawing.Size(136, 17)
        Me.lblTankOverfillProtectionType.TabIndex = 152
        Me.lblTankOverfillProtectionType.Text = "Overfill Protection Type:"
        Me.lblTankOverfillProtectionType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkEmergencyPower
        '
        Me.chkEmergencyPower.Location = New System.Drawing.Point(517, 8)
        Me.chkEmergencyPower.Name = "chkEmergencyPower"
        Me.chkEmergencyPower.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkEmergencyPower.Size = New System.Drawing.Size(207, 24)
        Me.chkEmergencyPower.TabIndex = 6
        Me.chkEmergencyPower.Text = "Tank used for emergency power"
        Me.chkEmergencyPower.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlTankMaterialHead
        '
        Me.pnlTankMaterialHead.Controls.Add(Me.lblTankMaterialHead)
        Me.pnlTankMaterialHead.Controls.Add(Me.lblTankMaterialDisplay)
        Me.pnlTankMaterialHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankMaterialHead.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlTankMaterialHead.Location = New System.Drawing.Point(2, 416)
        Me.pnlTankMaterialHead.Name = "pnlTankMaterialHead"
        Me.pnlTankMaterialHead.Size = New System.Drawing.Size(852, 24)
        Me.pnlTankMaterialHead.TabIndex = 181
        '
        'lblTankMaterialHead
        '
        Me.lblTankMaterialHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTankMaterialHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTankMaterialHead.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankMaterialHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTankMaterialHead.Location = New System.Drawing.Point(16, 0)
        Me.lblTankMaterialHead.Name = "lblTankMaterialHead"
        Me.lblTankMaterialHead.Size = New System.Drawing.Size(836, 24)
        Me.lblTankMaterialHead.TabIndex = 1
        Me.lblTankMaterialHead.Text = "Tank Material of Construction"
        Me.lblTankMaterialHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTankMaterialDisplay
        '
        Me.lblTankMaterialDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTankMaterialDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTankMaterialDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankMaterialDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblTankMaterialDisplay.Name = "lblTankMaterialDisplay"
        Me.lblTankMaterialDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblTankMaterialDisplay.TabIndex = 0
        Me.lblTankMaterialDisplay.Text = "-"
        Me.lblTankMaterialDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlTankTotalCapacity
        '
        Me.pnlTankTotalCapacity.Controls.Add(Me.txtTankCapacity)
        Me.pnlTankTotalCapacity.Controls.Add(Me.txtTankCompartmentNumber)
        Me.pnlTankTotalCapacity.Controls.Add(Me.pnlNonCompProperties)
        Me.pnlTankTotalCapacity.Controls.Add(Me.dGridCompartments)
        Me.pnlTankTotalCapacity.Controls.Add(Me.lblTankType)
        Me.pnlTankTotalCapacity.Controls.Add(Me.cmbTankType)
        Me.pnlTankTotalCapacity.Controls.Add(Me.cmbTankCompCercla)
        Me.pnlTankTotalCapacity.Controls.Add(Me.cmbTankCompSubstance)
        Me.pnlTankTotalCapacity.Controls.Add(Me.lblTankCapacity)
        Me.pnlTankTotalCapacity.Controls.Add(Me.chkTankCompartment)
        Me.pnlTankTotalCapacity.Controls.Add(Me.lblTankCompartmentNumber)
        Me.pnlTankTotalCapacity.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankTotalCapacity.Location = New System.Drawing.Point(2, 200)
        Me.pnlTankTotalCapacity.Name = "pnlTankTotalCapacity"
        Me.pnlTankTotalCapacity.Size = New System.Drawing.Size(852, 216)
        Me.pnlTankTotalCapacity.TabIndex = 2
        '
        'txtTankCapacity
        '
        Me.txtTankCapacity.Location = New System.Drawing.Point(248, 9)
        Me.txtTankCapacity.Name = "txtTankCapacity"
        Me.txtTankCapacity.Size = New System.Drawing.Size(48, 16)
        Me.txtTankCapacity.TabIndex = 197
        '
        'txtTankCompartmentNumber
        '
        Me.txtTankCompartmentNumber.Location = New System.Drawing.Point(512, 9)
        Me.txtTankCompartmentNumber.Name = "txtTankCompartmentNumber"
        Me.txtTankCompartmentNumber.Size = New System.Drawing.Size(48, 16)
        Me.txtTankCompartmentNumber.TabIndex = 196
        '
        'pnlNonCompProperties
        '
        Me.pnlNonCompProperties.BackColor = System.Drawing.SystemColors.Control
        Me.pnlNonCompProperties.Controls.Add(Me.lblCERCLAtt)
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankManifoldValue)
        Me.pnlNonCompProperties.Controls.Add(Me.cmbTankCercla)
        Me.pnlNonCompProperties.Controls.Add(Me.cmbTankFuelType)
        Me.pnlNonCompProperties.Controls.Add(Me.cmbTanksubstance)
        Me.pnlNonCompProperties.Controls.Add(Me.lblNonCompTankCapacity)
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankCercla)
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankFuelType)
        Me.pnlNonCompProperties.Controls.Add(Me.txtNonCompTankCapacity)
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankSubstance)
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankManifold)
        Me.pnlNonCompProperties.Controls.Add(Me.cmbTankCerclaDesc)
        Me.pnlNonCompProperties.Location = New System.Drawing.Point(8, 48)
        Me.pnlNonCompProperties.Name = "pnlNonCompProperties"
        Me.pnlNonCompProperties.Size = New System.Drawing.Size(736, 112)
        Me.pnlNonCompProperties.TabIndex = 3
        '
        'lblCERCLAtt
        '
        Me.lblCERCLAtt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCERCLAtt.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCERCLAtt.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblCERCLAtt.Location = New System.Drawing.Point(634, 45)
        Me.lblCERCLAtt.Name = "lblCERCLAtt"
        Me.lblCERCLAtt.Size = New System.Drawing.Size(16, 23)
        Me.lblCERCLAtt.TabIndex = 220
        Me.lblCERCLAtt.Text = "i"
        Me.lblCERCLAtt.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblTankManifoldValue
        '
        Me.lblTankManifoldValue.AutoSize = True
        Me.lblTankManifoldValue.Location = New System.Drawing.Point(120, 80)
        Me.lblTankManifoldValue.Name = "lblTankManifoldValue"
        Me.lblTankManifoldValue.Size = New System.Drawing.Size(0, 17)
        Me.lblTankManifoldValue.TabIndex = 207
        '
        'cmbTankCercla
        '
        Me.cmbTankCercla.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankCercla.DropDownWidth = 400
        Me.cmbTankCercla.Location = New System.Drawing.Point(424, 45)
        Me.cmbTankCercla.Name = "cmbTankCercla"
        Me.cmbTankCercla.Size = New System.Drawing.Size(208, 23)
        Me.cmbTankCercla.TabIndex = 3
        '
        'cmbTankFuelType
        '
        Me.cmbTankFuelType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankFuelType.Location = New System.Drawing.Point(424, 16)
        Me.cmbTankFuelType.Name = "cmbTankFuelType"
        Me.cmbTankFuelType.Size = New System.Drawing.Size(208, 23)
        Me.cmbTankFuelType.TabIndex = 2
        '
        'cmbTanksubstance
        '
        Me.cmbTanksubstance.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTanksubstance.DropDownWidth = 250
        Me.cmbTanksubstance.Location = New System.Drawing.Point(120, 45)
        Me.cmbTanksubstance.Name = "cmbTanksubstance"
        Me.cmbTanksubstance.Size = New System.Drawing.Size(208, 23)
        Me.cmbTanksubstance.TabIndex = 1
        '
        'lblNonCompTankCapacity
        '
        Me.lblNonCompTankCapacity.Location = New System.Drawing.Point(24, 16)
        Me.lblNonCompTankCapacity.Name = "lblNonCompTankCapacity"
        Me.lblNonCompTankCapacity.Size = New System.Drawing.Size(88, 23)
        Me.lblNonCompTankCapacity.TabIndex = 197
        Me.lblNonCompTankCapacity.Text = "Tank Capacity:"
        '
        'lblTankCercla
        '
        Me.lblTankCercla.Location = New System.Drawing.Point(360, 46)
        Me.lblTankCercla.Name = "lblTankCercla"
        Me.lblTankCercla.Size = New System.Drawing.Size(56, 23)
        Me.lblTankCercla.TabIndex = 202
        Me.lblTankCercla.Text = "Cercla #:"
        '
        'lblTankFuelType
        '
        Me.lblTankFuelType.Location = New System.Drawing.Point(352, 16)
        Me.lblTankFuelType.Name = "lblTankFuelType"
        Me.lblTankFuelType.Size = New System.Drawing.Size(64, 23)
        Me.lblTankFuelType.TabIndex = 201
        Me.lblTankFuelType.Text = "Fuel Type: "
        '
        'txtNonCompTankCapacity
        '
        Me.txtNonCompTankCapacity.Location = New System.Drawing.Point(120, 16)
        Me.txtNonCompTankCapacity.Name = "txtNonCompTankCapacity"
        Me.txtNonCompTankCapacity.Size = New System.Drawing.Size(80, 21)
        Me.txtNonCompTankCapacity.TabIndex = 0
        Me.txtNonCompTankCapacity.Text = ""
        '
        'lblTankSubstance
        '
        Me.lblTankSubstance.Location = New System.Drawing.Point(40, 46)
        Me.lblTankSubstance.Name = "lblTankSubstance"
        Me.lblTankSubstance.Size = New System.Drawing.Size(72, 23)
        Me.lblTankSubstance.TabIndex = 203
        Me.lblTankSubstance.Text = "Substance: "
        '
        'lblTankManifold
        '
        Me.lblTankManifold.Location = New System.Drawing.Point(56, 80)
        Me.lblTankManifold.Name = "lblTankManifold"
        Me.lblTankManifold.Size = New System.Drawing.Size(56, 23)
        Me.lblTankManifold.TabIndex = 199
        Me.lblTankManifold.Text = "Manifold:"
        '
        'cmbTankCerclaDesc
        '
        Me.cmbTankCerclaDesc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankCerclaDesc.DropDownWidth = 400
        Me.cmbTankCerclaDesc.Location = New System.Drawing.Point(664, 44)
        Me.cmbTankCerclaDesc.Name = "cmbTankCerclaDesc"
        Me.cmbTankCerclaDesc.Size = New System.Drawing.Size(64, 23)
        Me.cmbTankCerclaDesc.TabIndex = 3
        Me.cmbTankCerclaDesc.Visible = False
        '
        'dGridCompartments
        '
        Me.dGridCompartments.Cursor = System.Windows.Forms.Cursors.Default
        Me.dGridCompartments.Location = New System.Drawing.Point(8, 48)
        Me.dGridCompartments.Name = "dGridCompartments"
        Me.dGridCompartments.Size = New System.Drawing.Size(816, 152)
        Me.dGridCompartments.TabIndex = 195
        '
        'lblTankType
        '
        Me.lblTankType.AutoSize = True
        Me.lblTankType.Location = New System.Drawing.Point(544, 208)
        Me.lblTankType.Name = "lblTankType"
        Me.lblTankType.Size = New System.Drawing.Size(66, 17)
        Me.lblTankType.TabIndex = 194
        Me.lblTankType.Text = "Tank Type:"
        Me.lblTankType.Visible = False
        '
        'cmbTankType
        '
        Me.cmbTankType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankType.DropDownWidth = 200
        Me.cmbTankType.Location = New System.Drawing.Point(616, 208)
        Me.cmbTankType.Name = "cmbTankType"
        Me.cmbTankType.Size = New System.Drawing.Size(96, 23)
        Me.cmbTankType.TabIndex = 193
        Me.cmbTankType.Visible = False
        '
        'cmbTankCompCercla
        '
        Me.cmbTankCompCercla.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankCompCercla.DropDownWidth = 400
        Me.cmbTankCompCercla.Items.AddRange(New Object() {"12345 - The unmentionable", "23456 - Poison Gas"})
        Me.cmbTankCompCercla.Location = New System.Drawing.Point(448, 32)
        Me.cmbTankCompCercla.Name = "cmbTankCompCercla"
        Me.cmbTankCompCercla.Size = New System.Drawing.Size(96, 23)
        Me.cmbTankCompCercla.TabIndex = 192
        Me.cmbTankCompCercla.Visible = False
        '
        'cmbTankCompSubstance
        '
        Me.cmbTankCompSubstance.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankCompSubstance.DropDownWidth = 200
        Me.cmbTankCompSubstance.Items.AddRange(New Object() {"Not Harmful", "Hazardous"})
        Me.cmbTankCompSubstance.Location = New System.Drawing.Point(560, 32)
        Me.cmbTankCompSubstance.Name = "cmbTankCompSubstance"
        Me.cmbTankCompSubstance.Size = New System.Drawing.Size(96, 23)
        Me.cmbTankCompSubstance.TabIndex = 191
        Me.cmbTankCompSubstance.Visible = False
        '
        'lblTankCapacity
        '
        Me.lblTankCapacity.AutoSize = True
        Me.lblTankCapacity.Location = New System.Drawing.Point(160, 8)
        Me.lblTankCapacity.Name = "lblTankCapacity"
        Me.lblTankCapacity.Size = New System.Drawing.Size(83, 17)
        Me.lblTankCapacity.TabIndex = 189
        Me.lblTankCapacity.Text = "Total Capacity"
        '
        'chkTankCompartment
        '
        Me.chkTankCompartment.CheckAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.chkTankCompartment.Location = New System.Drawing.Point(40, 8)
        Me.chkTankCompartment.Name = "chkTankCompartment"
        Me.chkTankCompartment.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkTankCompartment.Size = New System.Drawing.Size(104, 16)
        Me.chkTankCompartment.TabIndex = 0
        Me.chkTankCompartment.Text = ":Compartments"
        Me.chkTankCompartment.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblTankCompartmentNumber
        '
        Me.lblTankCompartmentNumber.AutoSize = True
        Me.lblTankCompartmentNumber.Location = New System.Drawing.Point(352, 8)
        Me.lblTankCompartmentNumber.Name = "lblTankCompartmentNumber"
        Me.lblTankCompartmentNumber.Size = New System.Drawing.Size(150, 17)
        Me.lblTankCompartmentNumber.TabIndex = 187
        Me.lblTankCompartmentNumber.Text = "Number of Compartments:"
        '
        'Panel8
        '
        Me.Panel8.Controls.Add(Me.pnllblTankTotalCapacity)
        Me.Panel8.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel8.Location = New System.Drawing.Point(2, 176)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(852, 24)
        Me.Panel8.TabIndex = 189
        '
        'pnllblTankTotalCapacity
        '
        Me.pnllblTankTotalCapacity.Controls.Add(Me.lblTankTotalCapcityCaption)
        Me.pnllblTankTotalCapacity.Controls.Add(Me.lblTankTotalCapacity)
        Me.pnllblTankTotalCapacity.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnllblTankTotalCapacity.Location = New System.Drawing.Point(0, 0)
        Me.pnllblTankTotalCapacity.Name = "pnllblTankTotalCapacity"
        Me.pnllblTankTotalCapacity.Size = New System.Drawing.Size(852, 24)
        Me.pnllblTankTotalCapacity.TabIndex = 186
        '
        'lblTankTotalCapcityCaption
        '
        Me.lblTankTotalCapcityCaption.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTankTotalCapcityCaption.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTankTotalCapcityCaption.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTankTotalCapcityCaption.Location = New System.Drawing.Point(16, 0)
        Me.lblTankTotalCapcityCaption.Name = "lblTankTotalCapcityCaption"
        Me.lblTankTotalCapcityCaption.Size = New System.Drawing.Size(836, 24)
        Me.lblTankTotalCapcityCaption.TabIndex = 1
        Me.lblTankTotalCapcityCaption.Text = "Estimated Total Capacity(gallons)"
        Me.lblTankTotalCapcityCaption.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTankTotalCapacity
        '
        Me.lblTankTotalCapacity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTankTotalCapacity.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTankTotalCapacity.Location = New System.Drawing.Point(0, 0)
        Me.lblTankTotalCapacity.Name = "lblTankTotalCapacity"
        Me.lblTankTotalCapacity.Size = New System.Drawing.Size(16, 24)
        Me.lblTankTotalCapacity.TabIndex = 0
        Me.lblTankTotalCapacity.Text = "-"
        Me.lblTankTotalCapacity.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlTankInstallation
        '
        Me.pnlTankInstallation.Controls.Add(Me.lblTankInstalledOn)
        Me.pnlTankInstallation.Controls.Add(Me.cmbTankManufacturer)
        Me.pnlTankInstallation.Controls.Add(Me.lblTankManufacturer)
        Me.pnlTankInstallation.Controls.Add(Me.dtPickDatePlacedInService)
        Me.pnlTankInstallation.Controls.Add(Me.lblTankPlacedInServiceOn)
        Me.pnlTankInstallation.Controls.Add(Me.dtPickTankInstalled)
        Me.pnlTankInstallation.Controls.Add(Me.dtPickPlannedInstallation)
        Me.pnlTankInstallation.Controls.Add(Me.lblTankInstallationPlanedFor)
        Me.pnlTankInstallation.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankInstallation.Location = New System.Drawing.Point(2, 90)
        Me.pnlTankInstallation.Name = "pnlTankInstallation"
        Me.pnlTankInstallation.Size = New System.Drawing.Size(852, 86)
        Me.pnlTankInstallation.TabIndex = 1
        '
        'lblTankInstalledOn
        '
        Me.lblTankInstalledOn.AutoSize = True
        Me.lblTankInstalledOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankInstalledOn.Location = New System.Drawing.Point(72, 8)
        Me.lblTankInstalledOn.Name = "lblTankInstalledOn"
        Me.lblTankInstalledOn.Size = New System.Drawing.Size(105, 17)
        Me.lblTankInstalledOn.TabIndex = 188
        Me.lblTankInstalledOn.Text = "Tank Installed On:"
        '
        'cmbTankManufacturer
        '
        Me.cmbTankManufacturer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankManufacturer.Location = New System.Drawing.Point(448, 8)
        Me.cmbTankManufacturer.Name = "cmbTankManufacturer"
        Me.cmbTankManufacturer.Size = New System.Drawing.Size(216, 23)
        Me.cmbTankManufacturer.TabIndex = 2
        '
        'lblTankManufacturer
        '
        Me.lblTankManufacturer.AutoSize = True
        Me.lblTankManufacturer.Location = New System.Drawing.Point(360, 8)
        Me.lblTankManufacturer.Name = "lblTankManufacturer"
        Me.lblTankManufacturer.Size = New System.Drawing.Size(84, 17)
        Me.lblTankManufacturer.TabIndex = 187
        Me.lblTankManufacturer.Text = "Manufacturer: "
        '
        'dtPickDatePlacedInService
        '
        Me.dtPickDatePlacedInService.Checked = False
        Me.dtPickDatePlacedInService.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickDatePlacedInService.Location = New System.Drawing.Point(176, 32)
        Me.dtPickDatePlacedInService.Name = "dtPickDatePlacedInService"
        Me.dtPickDatePlacedInService.ShowCheckBox = True
        Me.dtPickDatePlacedInService.Size = New System.Drawing.Size(104, 21)
        Me.dtPickDatePlacedInService.TabIndex = 1
        '
        'lblTankPlacedInServiceOn
        '
        Me.lblTankPlacedInServiceOn.AutoSize = True
        Me.lblTankPlacedInServiceOn.Location = New System.Drawing.Point(23, 32)
        Me.lblTankPlacedInServiceOn.Name = "lblTankPlacedInServiceOn"
        Me.lblTankPlacedInServiceOn.Size = New System.Drawing.Size(157, 17)
        Me.lblTankPlacedInServiceOn.TabIndex = 184
        Me.lblTankPlacedInServiceOn.Text = "Tank Placed in Service On: "
        '
        'dtPickTankInstalled
        '
        Me.dtPickTankInstalled.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickTankInstalled.Checked = False
        Me.dtPickTankInstalled.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickTankInstalled.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickTankInstalled.Location = New System.Drawing.Point(176, 8)
        Me.dtPickTankInstalled.Name = "dtPickTankInstalled"
        Me.dtPickTankInstalled.ShowCheckBox = True
        Me.dtPickTankInstalled.Size = New System.Drawing.Size(104, 21)
        Me.dtPickTankInstalled.TabIndex = 0
        Me.dtPickTankInstalled.Value = New Date(2004, 5, 27, 15, 6, 0, 371)
        '
        'dtPickPlannedInstallation
        '
        Me.dtPickPlannedInstallation.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickPlannedInstallation.Checked = False
        Me.dtPickPlannedInstallation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickPlannedInstallation.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPlannedInstallation.Location = New System.Drawing.Point(632, 32)
        Me.dtPickPlannedInstallation.Name = "dtPickPlannedInstallation"
        Me.dtPickPlannedInstallation.ShowCheckBox = True
        Me.dtPickPlannedInstallation.Size = New System.Drawing.Size(104, 21)
        Me.dtPickPlannedInstallation.TabIndex = 174
        Me.dtPickPlannedInstallation.Value = New Date(2004, 5, 27, 15, 6, 0, 442)
        Me.dtPickPlannedInstallation.Visible = False
        '
        'lblTankInstallationPlanedFor
        '
        Me.lblTankInstallationPlanedFor.AutoSize = True
        Me.lblTankInstallationPlanedFor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankInstallationPlanedFor.Location = New System.Drawing.Point(568, 40)
        Me.lblTankInstallationPlanedFor.Name = "lblTankInstallationPlanedFor"
        Me.lblTankInstallationPlanedFor.Size = New System.Drawing.Size(169, 17)
        Me.lblTankInstallationPlanedFor.TabIndex = 173
        Me.lblTankInstallationPlanedFor.Text = "Tank Installation Planned For:"
        Me.lblTankInstallationPlanedFor.Visible = False
        '
        'pnllblDateofInstallation
        '
        Me.pnllblDateofInstallation.Controls.Add(Me.lblDateofInstallation)
        Me.pnllblDateofInstallation.Controls.Add(Me.lblTankInstallation)
        Me.pnllblDateofInstallation.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnllblDateofInstallation.Location = New System.Drawing.Point(2, 66)
        Me.pnllblDateofInstallation.Name = "pnllblDateofInstallation"
        Me.pnllblDateofInstallation.Size = New System.Drawing.Size(852, 24)
        Me.pnllblDateofInstallation.TabIndex = 187
        '
        'lblDateofInstallation
        '
        Me.lblDateofInstallation.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblDateofInstallation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblDateofInstallation.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblDateofInstallation.Location = New System.Drawing.Point(16, 0)
        Me.lblDateofInstallation.Name = "lblDateofInstallation"
        Me.lblDateofInstallation.Size = New System.Drawing.Size(836, 24)
        Me.lblDateofInstallation.TabIndex = 1
        Me.lblDateofInstallation.Text = "Date of Installation(month/year)"
        '
        'lblTankInstallation
        '
        Me.lblTankInstallation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTankInstallation.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTankInstallation.Location = New System.Drawing.Point(0, 0)
        Me.lblTankInstallation.Name = "lblTankInstallation"
        Me.lblTankInstallation.Size = New System.Drawing.Size(16, 24)
        Me.lblTankInstallation.TabIndex = 0
        Me.lblTankInstallation.Text = " -"
        '
        'pnlTankDescriptionTop
        '
        Me.pnlTankDescriptionTop.Controls.Add(Me.chkTankProhibition)
        Me.pnlTankDescriptionTop.Controls.Add(Me.chkBoxReplacementTank)
        Me.pnlTankDescriptionTop.Controls.Add(Me.lblTankOriginalStatusID)
        Me.pnlTankDescriptionTop.Controls.Add(Me.cmbTankStatus)
        Me.pnlTankDescriptionTop.Controls.Add(Me.lblTankStatus)
        Me.pnlTankDescriptionTop.Controls.Add(Me.txtTankFacilityID)
        Me.pnlTankDescriptionTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankDescriptionTop.DockPadding.All = 2
        Me.pnlTankDescriptionTop.Location = New System.Drawing.Point(2, 26)
        Me.pnlTankDescriptionTop.Name = "pnlTankDescriptionTop"
        Me.pnlTankDescriptionTop.Size = New System.Drawing.Size(852, 40)
        Me.pnlTankDescriptionTop.TabIndex = 0
        '
        'chkTankProhibition
        '
        Me.chkTankProhibition.Location = New System.Drawing.Point(8, 8)
        Me.chkTankProhibition.Name = "chkTankProhibition"
        Me.chkTankProhibition.Size = New System.Drawing.Size(88, 24)
        Me.chkTankProhibition.TabIndex = 141
        Me.chkTankProhibition.Text = "Prohibition"
        '
        'chkBoxReplacementTank
        '
        Me.chkBoxReplacementTank.Location = New System.Drawing.Point(472, 8)
        Me.chkBoxReplacementTank.Name = "chkBoxReplacementTank"
        Me.chkBoxReplacementTank.Size = New System.Drawing.Size(128, 24)
        Me.chkBoxReplacementTank.TabIndex = 140
        Me.chkBoxReplacementTank.Text = "Replacement Tank"
        Me.chkBoxReplacementTank.Visible = False
        '
        'lblTankOriginalStatusID
        '
        Me.lblTankOriginalStatusID.Location = New System.Drawing.Point(600, 8)
        Me.lblTankOriginalStatusID.Name = "lblTankOriginalStatusID"
        Me.lblTankOriginalStatusID.TabIndex = 139
        Me.lblTankOriginalStatusID.Visible = False
        '
        'cmbTankStatus
        '
        Me.cmbTankStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankStatus.Location = New System.Drawing.Point(176, 8)
        Me.cmbTankStatus.Name = "cmbTankStatus"
        Me.cmbTankStatus.Size = New System.Drawing.Size(272, 23)
        Me.cmbTankStatus.TabIndex = 0
        '
        'lblTankStatus
        '
        Me.lblTankStatus.AutoSize = True
        Me.lblTankStatus.Location = New System.Drawing.Point(96, 11)
        Me.lblTankStatus.Name = "lblTankStatus"
        Me.lblTankStatus.Size = New System.Drawing.Size(74, 17)
        Me.lblTankStatus.TabIndex = 137
        Me.lblTankStatus.Text = "Tank Status:"
        '
        'txtTankFacilityID
        '
        Me.txtTankFacilityID.Location = New System.Drawing.Point(376, 16)
        Me.txtTankFacilityID.Name = "txtTankFacilityID"
        Me.txtTankFacilityID.Size = New System.Drawing.Size(56, 21)
        Me.txtTankFacilityID.TabIndex = 138
        Me.txtTankFacilityID.Text = ""
        Me.txtTankFacilityID.Visible = False
        '
        'pnlTankDescHead
        '
        Me.pnlTankDescHead.Controls.Add(Me.lblTankDescHead)
        Me.pnlTankDescHead.Controls.Add(Me.lblTankDescDisplay)
        Me.pnlTankDescHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankDescHead.Location = New System.Drawing.Point(2, 2)
        Me.pnlTankDescHead.Name = "pnlTankDescHead"
        Me.pnlTankDescHead.Size = New System.Drawing.Size(852, 24)
        Me.pnlTankDescHead.TabIndex = 9
        '
        'lblTankDescHead
        '
        Me.lblTankDescHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTankDescHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTankDescHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTankDescHead.Location = New System.Drawing.Point(16, 0)
        Me.lblTankDescHead.Name = "lblTankDescHead"
        Me.lblTankDescHead.Size = New System.Drawing.Size(836, 24)
        Me.lblTankDescHead.TabIndex = 3
        Me.lblTankDescHead.Text = "Status of Tank"
        Me.lblTankDescHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTankDescDisplay
        '
        Me.lblTankDescDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTankDescDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTankDescDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblTankDescDisplay.Name = "lblTankDescDisplay"
        Me.lblTankDescDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblTankDescDisplay.TabIndex = 2
        Me.lblTankDescDisplay.Text = "-"
        Me.lblTankDescDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlTankDetailHeader
        '
        Me.pnlTankDetailHeader.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.pnlTankDetailHeader.Controls.Add(Me.lblTankCountVal)
        Me.pnlTankDetailHeader.Controls.Add(Me.btnAddTank2)
        Me.pnlTankDetailHeader.Controls.Add(Me.lblTankIDValue)
        Me.pnlTankDetailHeader.Controls.Add(Me.lblTankID)
        Me.pnlTankDetailHeader.Controls.Add(Me.btnAddPipe)
        Me.pnlTankDetailHeader.Controls.Add(Me.btnAddExistingPipe)
        Me.pnlTankDetailHeader.Controls.Add(Me.btnDetachPipes)
        Me.pnlTankDetailHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankDetailHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlTankDetailHeader.Name = "pnlTankDetailHeader"
        Me.pnlTankDetailHeader.Size = New System.Drawing.Size(860, 24)
        Me.pnlTankDetailHeader.TabIndex = 0
        '
        'lblTankCountVal
        '
        Me.lblTankCountVal.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblTankCountVal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankCountVal.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblTankCountVal.Location = New System.Drawing.Point(88, 4)
        Me.lblTankCountVal.Name = "lblTankCountVal"
        Me.lblTankCountVal.Size = New System.Drawing.Size(39, 17)
        Me.lblTankCountVal.TabIndex = 207
        Me.lblTankCountVal.Text = "of ???"
        '
        'btnAddTank2
        '
        Me.btnAddTank2.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddTank2.Location = New System.Drawing.Point(168, 2)
        Me.btnAddTank2.Name = "btnAddTank2"
        Me.btnAddTank2.Size = New System.Drawing.Size(128, 22)
        Me.btnAddTank2.TabIndex = 0
        Me.btnAddTank2.Text = "Add Tank"
        Me.btnAddTank2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblTankIDValue
        '
        Me.lblTankIDValue.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblTankIDValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblTankIDValue.Location = New System.Drawing.Point(64, 4)
        Me.lblTankIDValue.Name = "lblTankIDValue"
        Me.lblTankIDValue.Size = New System.Drawing.Size(18, 17)
        Me.lblTankIDValue.TabIndex = 140
        Me.lblTankIDValue.Text = "00"
        Me.lblTankIDValue.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTankID
        '
        Me.lblTankID.AutoSize = True
        Me.lblTankID.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblTankID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankID.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblTankID.Location = New System.Drawing.Point(8, 4)
        Me.lblTankID.Name = "lblTankID"
        Me.lblTankID.Size = New System.Drawing.Size(47, 17)
        Me.lblTankID.TabIndex = 138
        Me.lblTankID.Text = "Tank #:"
        '
        'btnAddPipe
        '
        Me.btnAddPipe.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddPipe.Location = New System.Drawing.Point(296, 2)
        Me.btnAddPipe.Name = "btnAddPipe"
        Me.btnAddPipe.Size = New System.Drawing.Size(128, 22)
        Me.btnAddPipe.TabIndex = 1
        Me.btnAddPipe.Text = "Add New Pipe"
        '
        'btnAddExistingPipe
        '
        Me.btnAddExistingPipe.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddExistingPipe.Location = New System.Drawing.Point(424, 2)
        Me.btnAddExistingPipe.Name = "btnAddExistingPipe"
        Me.btnAddExistingPipe.Size = New System.Drawing.Size(128, 22)
        Me.btnAddExistingPipe.TabIndex = 2
        Me.btnAddExistingPipe.Text = "Add Existing Pipe"
        '
        'btnDetachPipes
        '
        Me.btnDetachPipes.BackColor = System.Drawing.SystemColors.Control
        Me.btnDetachPipes.Location = New System.Drawing.Point(552, 2)
        Me.btnDetachPipes.Name = "btnDetachPipes"
        Me.btnDetachPipes.Size = New System.Drawing.Size(128, 22)
        Me.btnDetachPipes.TabIndex = 2
        Me.btnDetachPipes.Text = "Detach Pipe(s)"
        '
        'pnlTankButtons
        '
        Me.pnlTankButtons.Controls.Add(Me.btnTankCancel)
        Me.pnlTankButtons.Controls.Add(Me.btnTankComments)
        Me.pnlTankButtons.Controls.Add(Me.btnToPipe)
        Me.pnlTankButtons.Controls.Add(Me.btnCopyTankProfileToNew)
        Me.pnlTankButtons.Controls.Add(Me.btnDeleteTank)
        Me.pnlTankButtons.Controls.Add(Me.btnTankSave)
        Me.pnlTankButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTankButtons.Location = New System.Drawing.Point(0, 393)
        Me.pnlTankButtons.Name = "pnlTankButtons"
        Me.pnlTankButtons.Size = New System.Drawing.Size(860, 40)
        Me.pnlTankButtons.TabIndex = 2
        '
        'btnTankCancel
        '
        Me.btnTankCancel.Enabled = False
        Me.btnTankCancel.Location = New System.Drawing.Point(222, 8)
        Me.btnTankCancel.Name = "btnTankCancel"
        Me.btnTankCancel.Size = New System.Drawing.Size(75, 26)
        Me.btnTankCancel.TabIndex = 1
        Me.btnTankCancel.Text = "Cancel"
        '
        'btnTankComments
        '
        Me.btnTankComments.Location = New System.Drawing.Point(640, 8)
        Me.btnTankComments.Name = "btnTankComments"
        Me.btnTankComments.Size = New System.Drawing.Size(88, 26)
        Me.btnTankComments.TabIndex = 5
        Me.btnTankComments.Text = "Comments"
        '
        'btnToPipe
        '
        Me.btnToPipe.Location = New System.Drawing.Point(544, 8)
        Me.btnToPipe.Name = "btnToPipe"
        Me.btnToPipe.Size = New System.Drawing.Size(88, 26)
        Me.btnToPipe.TabIndex = 4
        Me.btnToPipe.Text = "Go To Pipe"
        '
        'btnCopyTankProfileToNew
        '
        Me.btnCopyTankProfileToNew.Location = New System.Drawing.Point(304, 8)
        Me.btnCopyTankProfileToNew.Name = "btnCopyTankProfileToNew"
        Me.btnCopyTankProfileToNew.Size = New System.Drawing.Size(136, 26)
        Me.btnCopyTankProfileToNew.TabIndex = 2
        Me.btnCopyTankProfileToNew.Text = "Copy Tank Profile"
        '
        'btnDeleteTank
        '
        Me.btnDeleteTank.Enabled = False
        Me.btnDeleteTank.Location = New System.Drawing.Point(448, 8)
        Me.btnDeleteTank.Name = "btnDeleteTank"
        Me.btnDeleteTank.Size = New System.Drawing.Size(88, 26)
        Me.btnDeleteTank.TabIndex = 3
        Me.btnDeleteTank.Text = "Delete Tank"
        '
        'btnTankSave
        '
        Me.btnTankSave.Enabled = False
        Me.btnTankSave.Location = New System.Drawing.Point(128, 8)
        Me.btnTankSave.Name = "btnTankSave"
        Me.btnTankSave.Size = New System.Drawing.Size(88, 26)
        Me.btnTankSave.TabIndex = 0
        Me.btnTankSave.Text = "Save Tank"
        '
        'pnlTankDetailMainDisplay
        '
        Me.pnlTankDetailMainDisplay.Controls.Add(Me.btnExpandTP2)
        Me.pnlTankDetailMainDisplay.Controls.Add(Me.lnkLblNextTank)
        Me.pnlTankDetailMainDisplay.Controls.Add(Me.lnkLblPrevTank)
        Me.pnlTankDetailMainDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankDetailMainDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlTankDetailMainDisplay.Name = "pnlTankDetailMainDisplay"
        Me.pnlTankDetailMainDisplay.Size = New System.Drawing.Size(960, 32)
        Me.pnlTankDetailMainDisplay.TabIndex = 1
        '
        'btnExpandTP2
        '
        Me.btnExpandTP2.Location = New System.Drawing.Point(7, 5)
        Me.btnExpandTP2.Name = "btnExpandTP2"
        Me.btnExpandTP2.Size = New System.Drawing.Size(96, 24)
        Me.btnExpandTP2.TabIndex = 147
        Me.btnExpandTP2.Text = "Collapse All"
        Me.btnExpandTP2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lnkLblNextTank
        '
        Me.lnkLblNextTank.AutoSize = True
        Me.lnkLblNextTank.Enabled = False
        Me.lnkLblNextTank.Location = New System.Drawing.Point(808, 8)
        Me.lnkLblNextTank.Name = "lnkLblNextTank"
        Me.lnkLblNextTank.Size = New System.Drawing.Size(48, 17)
        Me.lnkLblNextTank.TabIndex = 141
        Me.lnkLblNextTank.TabStop = True
        Me.lnkLblNextTank.Text = "Next >>"
        '
        'lnkLblPrevTank
        '
        Me.lnkLblPrevTank.AutoSize = True
        Me.lnkLblPrevTank.Enabled = False
        Me.lnkLblPrevTank.Location = New System.Drawing.Point(728, 8)
        Me.lnkLblPrevTank.Name = "lnkLblPrevTank"
        Me.lnkLblPrevTank.Size = New System.Drawing.Size(70, 17)
        Me.lnkLblPrevTank.TabIndex = 140
        Me.lnkLblPrevTank.TabStop = True
        Me.lnkLblPrevTank.Text = "<< Previous"
        '
        'pnlTankCount2
        '
        Me.pnlTankCount2.Controls.Add(Me.lblTotalNoOfTanksValue2)
        Me.pnlTankCount2.Controls.Add(Me.Label3)
        Me.pnlTankCount2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTankCount2.Location = New System.Drawing.Point(0, 888)
        Me.pnlTankCount2.Name = "pnlTankCount2"
        Me.pnlTankCount2.Size = New System.Drawing.Size(960, 24)
        Me.pnlTankCount2.TabIndex = 97
        Me.pnlTankCount2.Visible = False
        '
        'lblTotalNoOfTanksValue2
        '
        Me.lblTotalNoOfTanksValue2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfTanksValue2.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfTanksValue2.Location = New System.Drawing.Point(200, 0)
        Me.lblTotalNoOfTanksValue2.Name = "lblTotalNoOfTanksValue2"
        Me.lblTotalNoOfTanksValue2.Size = New System.Drawing.Size(48, 24)
        Me.lblTotalNoOfTanksValue2.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label3.Location = New System.Drawing.Point(0, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(200, 24)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Number of Tanks at this Location:"
        '
        'tbPageSummary
        '
        Me.tbPageSummary.AutoScroll = True
        Me.tbPageSummary.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageSummary.Controls.Add(Me.pnlOwnerSummaryDetails)
        Me.tbPageSummary.Controls.Add(Me.Panel12)
        Me.tbPageSummary.Controls.Add(Me.pnlOwnerSummaryHeader)
        Me.tbPageSummary.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPageSummary.Location = New System.Drawing.Point(4, 22)
        Me.tbPageSummary.Name = "tbPageSummary"
        Me.tbPageSummary.Size = New System.Drawing.Size(964, 636)
        Me.tbPageSummary.TabIndex = 0
        Me.tbPageSummary.Text = "Owner Summary"
        Me.tbPageSummary.Visible = False
        '
        'pnlOwnerSummaryDetails
        '
        Me.pnlOwnerSummaryDetails.Controls.Add(Me.UCOwnerSummary)
        Me.pnlOwnerSummaryDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerSummaryDetails.Location = New System.Drawing.Point(0, 16)
        Me.pnlOwnerSummaryDetails.Name = "pnlOwnerSummaryDetails"
        Me.pnlOwnerSummaryDetails.Size = New System.Drawing.Size(952, 616)
        Me.pnlOwnerSummaryDetails.TabIndex = 7
        '
        'UCOwnerSummary
        '
        Me.UCOwnerSummary.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerSummary.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerSummary.Name = "UCOwnerSummary"
        Me.UCOwnerSummary.Size = New System.Drawing.Size(952, 616)
        Me.UCOwnerSummary.TabIndex = 0
        '
        'Panel12
        '
        Me.Panel12.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel12.DockPadding.Left = 10
        Me.Panel12.Location = New System.Drawing.Point(952, 16)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(8, 616)
        Me.Panel12.TabIndex = 6
        '
        'pnlOwnerSummaryHeader
        '
        Me.pnlOwnerSummaryHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerSummaryHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlOwnerSummaryHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerSummaryHeader.Name = "pnlOwnerSummaryHeader"
        Me.pnlOwnerSummaryHeader.Size = New System.Drawing.Size(960, 16)
        Me.pnlOwnerSummaryHeader.TabIndex = 2
        '
        'ctxMenuTankCompartment
        '
        Me.ctxMenuTankCompartment.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddCompPipe, Me.mnuShowCompPipes})
        '
        'mnuAddCompPipe
        '
        Me.mnuAddCompPipe.Index = 0
        Me.mnuAddCompPipe.Text = "Add Pipe"
        '
        'mnuShowCompPipes
        '
        Me.mnuShowCompPipes.Index = 1
        Me.mnuShowCompPipes.Text = "Show Pipes"
        '
        'ctxMenuTankPipe
        '
        Me.ctxMenuTankPipe.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MI_DeleteTankPipe, Me.MenuItem2, Me.MenuItem1, Me.MenuItem3})
        '
        'MI_DeleteTankPipe
        '
        Me.MI_DeleteTankPipe.Index = 0
        Me.MI_DeleteTankPipe.Text = "Delete Checked Tanks && Pipes"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MI_NewTank1, Me.MI_EditTank1, Me.LI_AttachPipes1, Me.MI_CopyTank1, Me.LI_AddTankCompartment1})
        Me.MenuItem2.Text = "Manage Tanks"
        '
        'MI_NewTank1
        '
        Me.MI_NewTank1.Index = 0
        Me.MI_NewTank1.Text = "New Tank"
        '
        'MI_EditTank1
        '
        Me.MI_EditTank1.Index = 1
        Me.MI_EditTank1.Text = "Edit Selected Tank"
        '
        'LI_AttachPipes1
        '
        Me.LI_AttachPipes1.Index = 2
        Me.LI_AttachPipes1.Text = "Attach Pipes to Selected Tank"
        '
        'MI_CopyTank1
        '
        Me.MI_CopyTank1.Index = 3
        Me.MI_CopyTank1.Text = "Copy Selected Tank to New"
        '
        'LI_AddTankCompartment1
        '
        Me.LI_AddTankCompartment1.Index = 4
        Me.LI_AddTankCompartment1.Text = "Add Compartments to Selected Tank"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem4, Me.MI_EditPipe, Me.MI_CopyPipe, Me.MI_DetachPipes})
        Me.MenuItem1.Text = "Manage Pipes"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 0
        Me.MenuItem4.Text = "New Pipe"
        '
        'MI_EditPipe
        '
        Me.MI_EditPipe.Index = 1
        Me.MI_EditPipe.Text = "Edit Selected Pipe"
        '
        'MI_CopyPipe
        '
        Me.MI_CopyPipe.Index = 2
        Me.MI_CopyPipe.Text = "Copy Selected Pipe to New"
        '
        'MI_DetachPipes
        '
        Me.MI_DetachPipes.Index = 3
        Me.MI_DetachPipes.Text = "Detach Checked Pipes"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 3
        Me.MenuItem3.Text = ""
        '
        'ctxMenuTank
        '
        Me.ctxMenuTank.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MI_NewTank, Me.MI_EditTank, Me.MI_CopyTank, Me.LI_AddTankCompartment, Me.MI_DeleteTank, Me.MI_AttachPipes})
        '
        'MI_NewTank
        '
        Me.MI_NewTank.Index = 0
        Me.MI_NewTank.Text = "New Tank"
        '
        'MI_EditTank
        '
        Me.MI_EditTank.Index = 1
        Me.MI_EditTank.Text = "Edit Selected Tank"
        '
        'MI_CopyTank
        '
        Me.MI_CopyTank.Index = 2
        Me.MI_CopyTank.Text = "Copy Selected Tank To New"
        '
        'LI_AddTankCompartment
        '
        Me.LI_AddTankCompartment.Index = 3
        Me.LI_AddTankCompartment.Text = "Add Compartments to Selected Tank"
        '
        'MI_DeleteTank
        '
        Me.MI_DeleteTank.Index = 4
        Me.MI_DeleteTank.Text = "Deleted Checked Tanks"
        '
        'MI_AttachPipes
        '
        Me.MI_AttachPipes.Index = 5
        Me.MI_AttachPipes.Text = "Attach Pipes to Selected Tank"
        '
        'TextBox1
        '
        Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Left
        Me.TextBox1.Location = New System.Drawing.Point(226, 0)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(638, 20)
        Me.TextBox1.TabIndex = 3
        Me.TextBox1.Text = ""
        '
        'DateTimePicker8
        '
        Me.DateTimePicker8.Dock = System.Windows.Forms.DockStyle.Left
        Me.DateTimePicker8.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker8.Location = New System.Drawing.Point(121, 0)
        Me.DateTimePicker8.Name = "DateTimePicker8"
        Me.DateTimePicker8.Size = New System.Drawing.Size(105, 20)
        Me.DateTimePicker8.TabIndex = 2
        Me.DateTimePicker8.Value = New Date(2004, 5, 17, 9, 11, 37, 20)
        '
        'DateTimePicker9
        '
        Me.DateTimePicker9.Dock = System.Windows.Forms.DockStyle.Left
        Me.DateTimePicker9.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker9.Location = New System.Drawing.Point(17, 0)
        Me.DateTimePicker9.Name = "DateTimePicker9"
        Me.DateTimePicker9.Size = New System.Drawing.Size(104, 20)
        Me.DateTimePicker9.TabIndex = 1
        Me.DateTimePicker9.Value = New Date(2004, 5, 17, 9, 11, 37, 61)
        '
        'CheckBox10
        '
        Me.CheckBox10.Dock = System.Windows.Forms.DockStyle.Left
        Me.CheckBox10.Location = New System.Drawing.Point(0, 0)
        Me.CheckBox10.Name = "CheckBox10"
        Me.CheckBox10.Size = New System.Drawing.Size(17, 24)
        Me.CheckBox10.TabIndex = 0
        Me.CheckBox10.Text = "CheckBox7"
        '
        'Panel16
        '
        Me.Panel16.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel16.Location = New System.Drawing.Point(0, 72)
        Me.Panel16.Name = "Panel16"
        Me.Panel16.Size = New System.Drawing.Size(884, 24)
        Me.Panel16.TabIndex = 3
        '
        'TextBox2
        '
        Me.TextBox2.Dock = System.Windows.Forms.DockStyle.Left
        Me.TextBox2.Location = New System.Drawing.Point(226, 0)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(638, 20)
        Me.TextBox2.TabIndex = 3
        Me.TextBox2.Text = ""
        '
        'DateTimePicker10
        '
        Me.DateTimePicker10.Dock = System.Windows.Forms.DockStyle.Left
        Me.DateTimePicker10.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker10.Location = New System.Drawing.Point(121, 0)
        Me.DateTimePicker10.Name = "DateTimePicker10"
        Me.DateTimePicker10.Size = New System.Drawing.Size(105, 20)
        Me.DateTimePicker10.TabIndex = 2
        Me.DateTimePicker10.Value = New Date(2004, 5, 17, 9, 11, 37, 191)
        '
        'DateTimePicker11
        '
        Me.DateTimePicker11.Dock = System.Windows.Forms.DockStyle.Left
        Me.DateTimePicker11.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker11.Location = New System.Drawing.Point(17, 0)
        Me.DateTimePicker11.Name = "DateTimePicker11"
        Me.DateTimePicker11.Size = New System.Drawing.Size(104, 20)
        Me.DateTimePicker11.TabIndex = 1
        Me.DateTimePicker11.Value = New Date(2004, 5, 17, 9, 11, 37, 211)
        '
        'CheckBox11
        '
        Me.CheckBox11.Dock = System.Windows.Forms.DockStyle.Left
        Me.CheckBox11.Location = New System.Drawing.Point(0, 0)
        Me.CheckBox11.Name = "CheckBox11"
        Me.CheckBox11.Size = New System.Drawing.Size(17, 24)
        Me.CheckBox11.TabIndex = 0
        Me.CheckBox11.Text = "CheckBox7"
        '
        'Panel17
        '
        Me.Panel17.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel17.Location = New System.Drawing.Point(0, 48)
        Me.Panel17.Name = "Panel17"
        Me.Panel17.Size = New System.Drawing.Size(884, 24)
        Me.Panel17.TabIndex = 2
        '
        'TextBox3
        '
        Me.TextBox3.Dock = System.Windows.Forms.DockStyle.Left
        Me.TextBox3.Location = New System.Drawing.Point(226, 0)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(638, 20)
        Me.TextBox3.TabIndex = 3
        Me.TextBox3.Text = ""
        '
        'DateTimePicker12
        '
        Me.DateTimePicker12.Dock = System.Windows.Forms.DockStyle.Left
        Me.DateTimePicker12.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker12.Location = New System.Drawing.Point(121, 0)
        Me.DateTimePicker12.Name = "DateTimePicker12"
        Me.DateTimePicker12.Size = New System.Drawing.Size(105, 20)
        Me.DateTimePicker12.TabIndex = 2
        Me.DateTimePicker12.Value = New Date(2004, 5, 17, 9, 11, 37, 301)
        '
        'DateTimePicker13
        '
        Me.DateTimePicker13.Dock = System.Windows.Forms.DockStyle.Left
        Me.DateTimePicker13.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker13.Location = New System.Drawing.Point(17, 0)
        Me.DateTimePicker13.Name = "DateTimePicker13"
        Me.DateTimePicker13.Size = New System.Drawing.Size(104, 20)
        Me.DateTimePicker13.TabIndex = 1
        Me.DateTimePicker13.Value = New Date(2004, 5, 17, 9, 11, 37, 321)
        '
        'CheckBox12
        '
        Me.CheckBox12.Dock = System.Windows.Forms.DockStyle.Left
        Me.CheckBox12.Location = New System.Drawing.Point(0, 0)
        Me.CheckBox12.Name = "CheckBox12"
        Me.CheckBox12.Size = New System.Drawing.Size(17, 24)
        Me.CheckBox12.TabIndex = 0
        Me.CheckBox12.Text = "CheckBox7"
        '
        'Panel18
        '
        Me.Panel18.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel18.Location = New System.Drawing.Point(0, 24)
        Me.Panel18.Name = "Panel18"
        Me.Panel18.Size = New System.Drawing.Size(884, 24)
        Me.Panel18.TabIndex = 1
        '
        'DateTimePicker14
        '
        Me.DateTimePicker14.Dock = System.Windows.Forms.DockStyle.Left
        Me.DateTimePicker14.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker14.Location = New System.Drawing.Point(121, 0)
        Me.DateTimePicker14.Name = "DateTimePicker14"
        Me.DateTimePicker14.Size = New System.Drawing.Size(105, 20)
        Me.DateTimePicker14.TabIndex = 2
        Me.DateTimePicker14.Value = New Date(2004, 5, 17, 9, 11, 37, 401)
        '
        'DateTimePicker15
        '
        Me.DateTimePicker15.Dock = System.Windows.Forms.DockStyle.Left
        Me.DateTimePicker15.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.DateTimePicker15.Location = New System.Drawing.Point(17, 0)
        Me.DateTimePicker15.Name = "DateTimePicker15"
        Me.DateTimePicker15.Size = New System.Drawing.Size(104, 20)
        Me.DateTimePicker15.TabIndex = 1
        Me.DateTimePicker15.Value = New Date(2004, 5, 17, 9, 11, 37, 421)
        '
        'CheckBox13
        '
        Me.CheckBox13.Dock = System.Windows.Forms.DockStyle.Left
        Me.CheckBox13.Location = New System.Drawing.Point(0, 0)
        Me.CheckBox13.Name = "CheckBox13"
        Me.CheckBox13.Size = New System.Drawing.Size(17, 24)
        Me.CheckBox13.TabIndex = 0
        Me.CheckBox13.Text = "CheckBox7"
        '
        'Panel19
        '
        Me.Panel19.Controls.Add(Me.TextBox4)
        Me.Panel19.Controls.Add(Me.DateTimePicker14)
        Me.Panel19.Controls.Add(Me.DateTimePicker15)
        Me.Panel19.Controls.Add(Me.CheckBox13)
        Me.Panel19.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel19.Location = New System.Drawing.Point(0, 0)
        Me.Panel19.Name = "Panel19"
        Me.Panel19.Size = New System.Drawing.Size(884, 24)
        Me.Panel19.TabIndex = 0
        '
        'TextBox4
        '
        Me.TextBox4.Dock = System.Windows.Forms.DockStyle.Left
        Me.TextBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox4.Location = New System.Drawing.Point(226, 0)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(638, 21)
        Me.TextBox4.TabIndex = 3
        Me.TextBox4.Text = "This site is under Inspection and will soon be determined"
        '
        'Button6
        '
        Me.Button6.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button6.Location = New System.Drawing.Point(224, 0)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(640, 21)
        Me.Button6.TabIndex = 3
        Me.Button6.Text = "Comments"
        '
        'Button7
        '
        Me.Button7.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button7.Location = New System.Drawing.Point(120, 0)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(104, 21)
        Me.Button7.TabIndex = 2
        Me.Button7.Text = "Due Date"
        '
        'Button8
        '
        Me.Button8.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button8.Location = New System.Drawing.Point(19, 0)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(101, 21)
        Me.Button8.TabIndex = 1
        Me.Button8.Text = "Date"
        '
        'Button9
        '
        Me.Button9.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button9.Location = New System.Drawing.Point(0, 0)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(19, 21)
        Me.Button9.TabIndex = 0
        '
        'btnTankCompCol8
        '
        Me.btnTankCompCol8.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnTankCompCol8.Location = New System.Drawing.Point(534, 0)
        Me.btnTankCompCol8.Name = "btnTankCompCol8"
        Me.btnTankCompCol8.Size = New System.Drawing.Size(75, 20)
        Me.btnTankCompCol8.TabIndex = 7
        Me.btnTankCompCol8.Text = "Capacity"
        '
        'btnTankCompCol7
        '
        Me.btnTankCompCol7.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnTankCompCol7.Location = New System.Drawing.Point(399, 0)
        Me.btnTankCompCol7.Name = "btnTankCompCol7"
        Me.btnTankCompCol7.Size = New System.Drawing.Size(135, 20)
        Me.btnTankCompCol7.TabIndex = 6
        Me.btnTankCompCol7.Text = "Description"
        '
        'btnTankCompCol6
        '
        Me.btnTankCompCol6.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnTankCompCol6.Location = New System.Drawing.Point(324, 0)
        Me.btnTankCompCol6.Name = "btnTankCompCol6"
        Me.btnTankCompCol6.Size = New System.Drawing.Size(75, 20)
        Me.btnTankCompCol6.TabIndex = 5
        Me.btnTankCompCol6.Text = "Cercla.No"
        '
        'btnTankCompCol5
        '
        Me.btnTankCompCol5.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnTankCompCol5.Location = New System.Drawing.Point(249, 0)
        Me.btnTankCompCol5.Name = "btnTankCompCol5"
        Me.btnTankCompCol5.Size = New System.Drawing.Size(75, 20)
        Me.btnTankCompCol5.TabIndex = 4
        Me.btnTankCompCol5.Text = "Fuel Type"
        '
        'btnTankCompCol4
        '
        Me.btnTankCompCol4.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnTankCompCol4.Location = New System.Drawing.Point(174, 0)
        Me.btnTankCompCol4.Name = "btnTankCompCol4"
        Me.btnTankCompCol4.Size = New System.Drawing.Size(75, 20)
        Me.btnTankCompCol4.TabIndex = 3
        Me.btnTankCompCol4.Text = "Substance"
        '
        'btnTankCompCol3
        '
        Me.btnTankCompCol3.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnTankCompCol3.Location = New System.Drawing.Point(99, 0)
        Me.btnTankCompCol3.Name = "btnTankCompCol3"
        Me.btnTankCompCol3.Size = New System.Drawing.Size(75, 20)
        Me.btnTankCompCol3.TabIndex = 2
        Me.btnTankCompCol3.Text = "No"
        '
        'btnTankCompCol2
        '
        Me.btnTankCompCol2.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnTankCompCol2.Location = New System.Drawing.Point(24, 0)
        Me.btnTankCompCol2.Name = "btnTankCompCol2"
        Me.btnTankCompCol2.Size = New System.Drawing.Size(75, 20)
        Me.btnTankCompCol2.TabIndex = 1
        Me.btnTankCompCol2.Text = "Pipe ID"
        '
        'btnTankCompCol1
        '
        Me.btnTankCompCol1.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnTankCompCol1.Location = New System.Drawing.Point(0, 0)
        Me.btnTankCompCol1.Name = "btnTankCompCol1"
        Me.btnTankCompCol1.Size = New System.Drawing.Size(24, 20)
        Me.btnTankCompCol1.TabIndex = 0
        '
        'pnlTankCompartmentHeader
        '
        Me.pnlTankCompartmentHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankCompartmentHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlTankCompartmentHeader.Name = "pnlTankCompartmentHeader"
        Me.pnlTankCompartmentHeader.Size = New System.Drawing.Size(618, 20)
        Me.pnlTankCompartmentHeader.TabIndex = 0
        '
        'lblNoOfPreviouslyOwnedFacilitiesValue
        '
        Me.lblNoOfPreviouslyOwnedFacilitiesValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfPreviouslyOwnedFacilitiesValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfPreviouslyOwnedFacilitiesValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoOfPreviouslyOwnedFacilitiesValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNoOfPreviouslyOwnedFacilitiesValue.Location = New System.Drawing.Point(100, 0)
        Me.lblNoOfPreviouslyOwnedFacilitiesValue.Name = "lblNoOfPreviouslyOwnedFacilitiesValue"
        Me.lblNoOfPreviouslyOwnedFacilitiesValue.Size = New System.Drawing.Size(36, 22)
        Me.lblNoOfPreviouslyOwnedFacilitiesValue.TabIndex = 1
        '
        'lblNoOfPreviouslyOwnedFacilities
        '
        Me.lblNoOfPreviouslyOwnedFacilities.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfPreviouslyOwnedFacilities.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfPreviouslyOwnedFacilities.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNoOfPreviouslyOwnedFacilities.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNoOfPreviouslyOwnedFacilities.Location = New System.Drawing.Point(0, 0)
        Me.lblNoOfPreviouslyOwnedFacilities.Name = "lblNoOfPreviouslyOwnedFacilities"
        Me.lblNoOfPreviouslyOwnedFacilities.Size = New System.Drawing.Size(100, 22)
        Me.lblNoOfPreviouslyOwnedFacilities.TabIndex = 0
        '
        'Registration
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(972, 694)
        Me.Controls.Add(Me.pnlMain)
        Me.Controls.Add(Me.pnlTop)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Registration"
        Me.Text = "Registration"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTop.ResumeLayout(False)
        Me.pnlMain.ResumeLayout(False)
        Me.tbCntrlRegistration.ResumeLayout(False)
        Me.tbPageOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerBottom.ResumeLayout(False)
        Me.tbCtrlOwner.ResumeLayout(False)
        Me.tbPageOwnerFacilities.ResumeLayout(False)
        Me.pnlOwnerFacilityBottom.ResumeLayout(False)
        Me.tbPageOwnerContactList.ResumeLayout(False)
        Me.pnlOwnerContactContainer.ResumeLayout(False)
        Me.pnlOwnerContactHeader.ResumeLayout(False)
        Me.pnlOwnerContactButtons.ResumeLayout(False)
        Me.tbPrevFacs.ResumeLayout(False)
        Me.pnlPrevOwnedFacsCount.ResumeLayout(False)
        Me.tbPageOwnerDocuments.ResumeLayout(False)
        Me.pnlOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerName.ResumeLayout(False)
        Me.pnlOwnerNameButton.ResumeLayout(False)
        Me.pnlOwnerOrg.ResumeLayout(False)
        Me.pnlPersonOrganization.ResumeLayout(False)
        Me.pnlOwnerPerson.ResumeLayout(False)
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerButtons.ResumeLayout(False)
        Me.tbPageFacilityDetail.ResumeLayout(False)
        Me.pnlFacilityBottom.ResumeLayout(False)
        Me.tabCtrlFacilityTankPipe.ResumeLayout(False)
        Me.tbPageTankPipe.ResumeLayout(False)
        Me.pnlTankPipe.ResumeLayout(False)
        CType(Me.dgPipesAndTanks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityTankPipeButton.ResumeLayout(False)
        Me.tbPageFacilityContactList.ResumeLayout(False)
        Me.pnlFacilityContactContainer.ResumeLayout(False)
        CType(Me.ugFacilityContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityContactHeader.ResumeLayout(False)
        Me.pnlFacilityContactBottom.ResumeLayout(False)
        Me.tpPreviouslyOwnedOwners.ResumeLayout(False)
        CType(Me.ugPreviouslyOwnedOwners, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.tbPageFacilityDocuments.ResumeLayout(False)
        Me.pnl_FacilityDetail.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageManageTank.ResumeLayout(False)
        CType(Me.dgPipesAndTanks2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbCntrlPipe.ResumeLayout(False)
        Me.tbPagePipeDetail.ResumeLayout(False)
        Me.pnlPipeDetail.ResumeLayout(False)
        Me.pnlPipeClosure.ResumeLayout(False)
        Me.Panel20.ResumeLayout(False)
        Me.pnlPipeInstallerOath.ResumeLayout(False)
        Me.Panel9.ResumeLayout(False)
        Me.pnlPipeRelease.ResumeLayout(False)
        Me.grpBxReleaseDetectionGroup2.ResumeLayout(False)
        Me.grpBxPipeReleaseDetectionGroup1.ResumeLayout(False)
        Me.Panel11.ResumeLayout(False)
        Me.pnlPipeType.ResumeLayout(False)
        Me.Panel15.ResumeLayout(False)
        Me.pnlPipeMaterial.ResumeLayout(False)
        Me.grpBxPipeTerminationAtDispenser.ResumeLayout(False)
        Me.grpBxPipeTerminationAtTank.ResumeLayout(False)
        Me.grpBxPipeContainmentSumpsLocation.ResumeLayout(False)
        Me.pnlPipeMaterialHead.ResumeLayout(False)
        Me.pnlPipeDateOfInstallation.ResumeLayout(False)
        Me.Panel13.ResumeLayout(False)
        Me.pnlPipeDescription.ResumeLayout(False)
        Me.pnlPipeDescHead.ResumeLayout(False)
        Me.pnlPipeButtons.ResumeLayout(False)
        Me.pnlPipeDetailHeader.ResumeLayout(False)
        Me.tbCntrlTank.ResumeLayout(False)
        Me.tbPageTankDetail.ResumeLayout(False)
        Me.pnlTankDetail.ResumeLayout(False)
        Me.pnlTankClosure.ResumeLayout(False)
        Me.pnllblTankClosure.ResumeLayout(False)
        Me.pnlTankInstallerOath.ResumeLayout(False)
        Me.pnlInstallerOath.ResumeLayout(False)
        Me.pnlTankRelease.ResumeLayout(False)
        Me.pnlReleaseDetection.ResumeLayout(False)
        Me.pnlTankMaterial.ResumeLayout(False)
        Me.pnlTankMaterialHead.ResumeLayout(False)
        Me.pnlTankTotalCapacity.ResumeLayout(False)
        Me.pnlNonCompProperties.ResumeLayout(False)
        CType(Me.dGridCompartments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel8.ResumeLayout(False)
        Me.pnllblTankTotalCapacity.ResumeLayout(False)
        Me.pnlTankInstallation.ResumeLayout(False)
        Me.pnllblDateofInstallation.ResumeLayout(False)
        Me.pnlTankDescriptionTop.ResumeLayout(False)
        Me.pnlTankDescHead.ResumeLayout(False)
        Me.pnlTankDetailHeader.ResumeLayout(False)
        Me.pnlTankButtons.ResumeLayout(False)
        Me.pnlTankDetailMainDisplay.ResumeLayout(False)
        Me.pnlTankCount2.ResumeLayout(False)
        Me.tbPageSummary.ResumeLayout(False)
        Me.pnlOwnerSummaryDetails.ResumeLayout(False)
        Me.Panel19.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "General"
    Private Sub ShowError(ByVal ex As Exception)
        Dim MyErr As New ErrorReport(ex)
        MyErr.ShowDialog()
    End Sub

    Private Sub ResetOwnerID()
        nOwnerID = 0
        'ResetFacilityID()
        ResetTankID()
        ResetCompartmentNumber()
        ResetPipeID()
    End Sub
    Private Sub ResetFacilityID()
        nFacilityID = 0
        ResetTankID()
        ResetCompartmentNumber()
        ResetPipeID()
    End Sub
    Private Sub ResetTankID()
        nTankID = 0
        ResetCompartmentNumber()
        ResetPipeID()
    End Sub
    Private Sub ResetCompartmentNumber()
        nCompartmentNumber = 0
        ResetPipeID()
    End Sub
    Private Sub ResetPipeID()
        nPipeID = 0
    End Sub

    Private Sub CheckRegHeader(ByVal ownID As Integer)
        Try
            If oReg Is Nothing Then
                oReg = New MUSTER.BusinessLogic.pRegistration
                oReg.RetrieveByOwnerID(nOwnerID)
            ElseIf oReg.OWNER_ID <> nOwnerID Then
                oReg.RetrieveByOwnerID(nOwnerID)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Friend Sub PutRegistrationActivity(ByVal EntityID As Integer, ByVal EntityType As Integer, ByVal Activity As String)
        Try
            Dim bolAddActivity As Boolean = True
            Dim bolShowMsg As Boolean = True

            CheckRegHeader(nOwnerID)

            If oReg.ID <= 0 Then
                oReg.OWNER_ID = nOwnerID
                oReg.DATE_STARTED = Now
                oReg.DATE_COMPLETED = dtNullDate
                oReg.COMPLETED = False
                oReg.Deleted = False
                oReg.Save()
            End If

            oReg.Activity.Col = oReg.Activity.RetrieveByRegID(oReg.ID)

            If oReg.Activity.Col.Count > 0 Then
                bolShowMsg = False
                For Each oRegActinfo In oReg.Activity.Col.Values
                    If oRegActinfo.EntityId = EntityID And _
                        oRegActinfo.EntityType = EntityType And _
                        oRegActinfo.ActivityDesc = Activity Then
                        bolAddActivity = False
                        Exit For
                    End If
                Next
            End If

            If bolAddActivity Then
                ' create cal entry
                Dim calDescText As String = String.Empty
                Select Case Activity
                    Case UIUtilsGen.ActivityTypes.AddOwner
                        calDescText = nOwnerID.ToString + " Incomplete Registration - New Owner Added"
                        'Case UIUtilsGen.ActivityTypes.TransferAcknowledgement
                        '    calDescText = nOwnerID.ToString + " Incomplete Registration - Facility Transfer Acknowledgement"
                    Case UIUtilsGen.ActivityTypes.TransferOwnership
                        calDescText = nOwnerID.ToString + " Incomplete Registration - Facility Transferred"
                    Case UIUtilsGen.ActivityTypes.UpComingInstall
                        calDescText = nOwnerID.ToString + " Incomplete Registration - Upcoming Install for Facility (" + nFacilityID.ToString + ")"
                    Case UIUtilsGen.ActivityTypes.SignatureRequired
                        calDescText = nOwnerID.ToString + " Incomplete Registration - Signature Required for Facility (" + nFacilityID.ToString + ")"
                    Case UIUtilsGen.ActivityTypes.TankStatusTOSI
                        calDescText = nOwnerID.ToString + " Incomplete Registration - Tank Status changed to TOSI in Facility (" + nFacilityID.ToString + ")"
                    Case UIUtilsGen.ActivityTypes.AddTank
                        calDescText = nOwnerID.ToString + " Incomplete Registration - New Tank Added for Facility (" + nFacilityID.ToString + ")"
                End Select
                MC.pCalendar.Add(New MUSTER.Info.CalendarInfo(0, Now, DateAdd(DateInterval.Day, 30, Now), 0, calDescText, MC.AppUser.ID, "SYSTEM", "", False, True, False, False, MC.AppUser.ID, Now, String.Empty, dtNullDate, EntityType, EntityID))
                MC.pCalendar.Save()

                MC.RefreshCalendarInfo()
                MC.LoadDueToMeCalendar()
                MC.LoadToDoCalendar()

                'Adding Registration Activity Detail
                oRegActinfo = New MUSTER.Info.RegistrationActivityInfo(0, _
                                                            oReg.ID, _
                                                            EntityType, _
                                                            EntityID, _
                                                            MC.AppUser.ID, _
                                                            Activity, _
                                                            False, _
                                                            Now(), _
                                                            MC.pCalendar.CalendarId)
                oReg.Activity.Add(oRegActinfo)
                oReg.Activity.Save()
            End If

            If bolShowMsg Then
                MsgBox("Placing owner in registration mode.", MsgBoxStyle.Information & MsgBoxStyle.OKOnly, "Registration Initiated")
            End If

            If oReg.Activity.Values.Count > 0 Then
                btnRegister.Visible = True
            Else
                btnRegister.Visible = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub DeleteRegistrationActivity(ByVal EntityID As Integer, ByVal EntityType As Integer, ByVal Activity As String)
        Try
            Dim bolDeleteActivity As Boolean = False

            CheckRegHeader(nOwnerID)

            If oReg.ID > 0 Then
                oReg.Activity.Col = oReg.Activity.RetrieveByRegID(oReg.ID)
                If oReg.Activity.Col.Count > 0 Then
                    For Each oRegActinfo In oReg.Activity.Col.Values
                        If oRegActinfo.EntityId = EntityID And _
                            oRegActinfo.EntityType = EntityType And _
                            oRegActinfo.ActivityDesc = Activity Then
                            oReg.Activity.RegActivityInfo = oRegActinfo
                            bolDeleteActivity = True
                            Exit For
                        End If
                    Next
                End If
            End If



            If bolDeleteActivity Then
                ' delete cal entry - taken care in trigger
                oReg.Activity.Processed = True
                oReg.Activity.Deleted = True
                oReg.Activity.Save()

                MC.RefreshCalendarInfo()
                MC.LoadDueToMeCalendar()
                MC.LoadToDoCalendar()

                oReg.Activity.Col.Remove(oReg.Activity.RegActionIndex)

                If oReg.Activity.Values.Count > 0 Then
                    btnRegister.Visible = True
                Else
                    btnRegister.Visible = False
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub DeleteRegistrationActivities(ByVal ownID As Integer)
        Try
            CheckRegHeader(ownID)
            If oReg.ID > 0 Then
                oReg.Activity.Col = oReg.Activity.RetrieveByRegID(oReg.ID)
                ' deleting cal entries taken care in trigger
                For Each oRegActinfo In oReg.Activity.Col.Values
                    oRegActinfo.Deleted = True
                Next
                oReg.Deleted = True
                oReg.Save()

                MC.RefreshCalendarInfo()
                MC.LoadDueToMeCalendar()
                MC.LoadToDoCalendar()

            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Public Sub CheckForRegistrationActivity()
        Try
            btnRegister.Visible = False

            CheckRegHeader(nOwnerID)

            '------ check the registration activities to see if there is an unprocessed 
            '       registration activity... if the Retrieve returns new info.. then there is no instance
            '       in the info object for this owner
            If oReg.ID > 0 Then
                oReg.Activity.Col = oReg.Activity.RetrieveByRegID(oReg.ID)
                If oReg.Activity.Col.Count > 0 Then
                    btnRegister.Visible = True
                End If
            End If

            MC.RefreshCalendarInfo()
            MC.LoadDueToMeCalendar()
            MC.LoadToDoCalendar()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function CheckCapDates(ByVal tank As Boolean) As Boolean
        Dim returnVal As Boolean = True
        Dim strMsg As String = String.Empty
        Dim dttemp, dtValidDate As Date
        Try
            If tank Then
                If dtPickCPLastTested.Enabled Then
                    dttemp = pTank.LastTCPDate
                    dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                    dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                    If Date.Compare(pTank.TCPInstallDate, dtValidDate) > 0 Then
                        dtValidDate = pTank.TCPInstallDate
                    End If
                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                        strMsg += "Tank CP Last Tested : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString + vbCrLf
                    End If
                End If
                If dtPickLastInteriorLinningInspection.Enabled Then
                    dttemp = pTank.LinedInteriorInspectDate
                    dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString

                    ' install = null 10 yrs
                    If Date.Compare(pTank.LinedInteriorInstallDate, dtNullDate) = 0 Then
                        dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                    Else ' if install is more than 15 yrs old, 5 yrs
                        ' first inspection = 10yrs, second and onwards = 5yrs
                        If Date.Compare(pTank.LinedInteriorInstallDate, DateAdd(DateInterval.Year, -15, Today.Date)) <= 0 Then
                            dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                        Else
                            dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                        End If
                    End If

                    'If Date.Compare(pTank.LinedInteriorInspectDateOriginal, dtNullDate) = 0 Or _
                    '    Date.Compare(pTank.LinedInteriorInspectDate, dtNullDate) = 0 Or _
                    '    Date.Compare(DateAdd(DateInterval.Year, 10, pTank.LinedInteriorInstallDate), _
                    '                    DateAdd(DateInterval.Year, 5, pTank.LinedInteriorInspectDate)) > 0 Then
                    '    dtValidDate = DateAdd(DateInterval.Year, -10, dtValidDate)
                    'Else
                    '    dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                    'End If
                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                    If Date.Compare(pTank.LinedInteriorInstallDate, dtValidDate) > 0 Then
                        dtValidDate = pTank.LinedInteriorInstallDate
                    End If
                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                        strMsg += "Last InteriorLinning Inspection : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString
                    End If
                End If
                If dtPickTankTightnessTest.Enabled Then
                    dttemp = pTank.TTTDate
                    dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                    dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                        strMsg += "Tank Tightness Test : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString
                    End If
                End If
            Else
                If dtPickPipeCPLastTest.Enabled Then
                    dttemp = pPipe.PipeCPTest
                    dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                    dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                    If Date.Compare(pPipe.PipeCPInstalledDate, dtValidDate) > 0 Then
                        dtValidDate = pPipe.PipeCPInstalledDate
                    End If
                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                        strMsg += "Pipe CP Last Tested : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString
                    End If
                End If
                If dtPickPipeTerminationCPLastTested.Enabled Then
                    dttemp = pPipe.TermCPLastTested
                    dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                    dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                    If Date.Compare(pPipe.TermCPInstalledDate, dtValidDate) > 0 Then
                        dtValidDate = pPipe.TermCPInstalledDate
                    End If
                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                        strMsg += "Termination CP Last Tested : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString
                    End If
                End If
                If dtPickPipeTightnessTest.Enabled Then
                    dttemp = pPipe.LTTDate
                    If pPipe.PipeTypeDesc = 268 Then
                        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                        dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
                        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                    ElseIf pPipe.PipeTypeDesc = 266 Then
                        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                        dtValidDate = DateAdd(DateInterval.Year, -1, dtValidDate)
                        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                    Else
                        dtPickPipeTightnessTest.Refresh()
                        Exit Function
                    End If
                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                        strMsg += "Last Pipe Tightness Test : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString
                    End If
                End If
                If dtPickPipeLeakDetectorTest.Enabled Then
                    dttemp = pPipe.ALLDTestDate
                    dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
                    dtValidDate = DateAdd(DateInterval.Year, -1, dtValidDate)
                    dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
                    If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
                        strMsg += "Automatic Line Leak Detector Test : " + dtValidDate.ToShortDateString + " to " + dtTodayPlus90Days.ToShortDateString
                    End If
                End If
            End If
            If strMsg <> String.Empty Then
                If MsgBox("Invalid dates entered for CAP." + vbCrLf + "Valid dates for" + vbCrLf + strMsg + vbCrLf + "Do you want to continue?", MsgBoxStyle.YesNo, "CAP Validation") = MsgBoxResult.No Then
                    returnVal = False
                Else
                    returnVal = True
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return returnVal
    End Function
    Private Sub EnableDateField(ByRef dtPick As DateTimePicker, ByVal bolValue As Boolean)
        dtPick.Enabled = bolValue
    End Sub
    Private Function GetPrevNext(ByVal sl As SortedList, ByVal getNext As Boolean, ByVal key As Integer) As Integer
        Try
            Dim retVal As Integer
            Dim index As Integer = sl.IndexOfKey(key)

            If getNext Then
                If sl.Count = 1 Then
                    index = -1
                ElseIf index = sl.Count - 1 Then
                    index = -1
                End If
                retVal = sl.GetByIndex(index + 1)
            Else
                If sl.Count = 1 Then
                    index = 1
                ElseIf index = 0 Then
                    index = sl.Count
                End If
                retVal = sl.GetByIndex(index - 1)
            End If
            Return retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function GetPrevNext(ByVal sl As SortedList, ByVal getNext As Boolean, ByVal key As String) As String
        Try
            Dim retVal As String
            Dim index As String = sl.IndexOfKey(key)

            If getNext Then
                If sl.Count = 1 Then
                    index = -1
                ElseIf index = sl.Count - 1 Then
                    index = -1
                End If
                retVal = sl.GetByIndex(index + 1)
            Else
                If sl.Count = 1 Then
                    index = 1
                ElseIf index = 0 Then
                    index = sl.Count
                End If
                retVal = sl.GetByIndex(index - 1)
            End If
            Return retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub ShowHideTankPipeScreen(ByVal bolTank As Boolean, ByVal bolPipe As Boolean)
        Try
            tbCntrlTank.Visible = bolTank
            tbCntrlPipe.Visible = bolPipe

            If bolTank Or bolPipe Then
                pnlTankCount2.Visible = False
                dgPipesAndTanks2.Visible = False
                btnExpandTP2.Visible = False
                pnlTankCount2.Visible = False

                lnkLblNextTank.Visible = True
                lnkLblPrevTank.Visible = True
                lnkLblNextTank.Enabled = True
                lnkLblPrevTank.Enabled = True
            Else
                pnlTankCount2.Visible = True
                dgPipesAndTanks2.Visible = True
                btnExpandTP2.Visible = True
                pnlTankCount2.Visible = True

                lnkLblNextTank.Visible = False
                lnkLblPrevTank.Visible = False
                lnkLblNextTank.Enabled = False
                lnkLblPrevTank.Enabled = False
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub ExpandCollapse(ByRef pnl As Panel, ByRef lbl As Label)
        pnl.Visible = Not pnl.Visible
        lbl.Text = IIf(pnl.Visible, "-", "+")
    End Sub
    Private Sub ExpandAll(ByVal bol As Boolean, ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid, ByRef btn As Button)
        If bol Then
            btn.Text = "Collapse All"
            ug.Rows.ExpandAll(True)
        Else
            btn.Text = "Expand All"
            ug.Rows.CollapseAll(True)
        End If
    End Sub

    Private Function CheckTabs() As Boolean
        Dim bolLoadingLocal As Boolean = bolLoading
        Dim tb As TabPage = tbCntrlRegistration.SelectedTab
        Dim retVal As Boolean = True
        Try
            bolLoading = True
            If tbCntrlRegistration.SelectedTab.Name <> tbCntrlRegistration.Tag Then
                If bolAddOwner Then
                    tbCntrlRegistration.SelectedTab = tbPageOwnerDetail
                    MsgBox("Please Save Owner First")
                    retVal = False
                ElseIf bolAddFacility And tbCntrlRegistration.SelectedTab.Name <> tbPageFacilityDetail.Name Then
                    tbCntrlRegistration.SelectedTab = tbPageFacilityDetail
                    If MsgBox("Do you want to save the new Facility?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        pOwn.Facilities.Remove(nFacilityID)
                        ResetFacilityID()
                        bolAddFacility = False
                        tbCntrlRegistration.SelectedTab = tbPageOwnerDetail
                    Else
                        btnFacilitySave.PerformClick()
                        retVal = False
                    End If
                ElseIf bolAddTank Then
                    If tbCntrlRegistration.Tag = tbCntrlTank.Name And nTankID < 0 Then
                        ShowHideTankPipeScreen(True, False)
                        tbCntrlRegistration.SelectedTab = tbPageManageTank
                        If MsgBox("Do you want to save the new Tank?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            pTank.Remove(nTankID)
                            ResetTankID()
                            bolAddTank = False
                            If tb.Name = tbPageManageTank.Name Then
                                ShowHideTankPipeScreen(False, False)
                            Else
                                tbCntrlRegistration.SelectedTab = tb
                            End If
                        Else
                            btnTankSave.PerformClick()
                            retVal = False
                        End If
                    End If
                ElseIf bolAddPipe Then
                    If tbCntrlRegistration.Tag = tbCntrlPipe.Name And nPipeID < 0 Then
                        ShowHideTankPipeScreen(False, True)
                        tbCntrlRegistration.SelectedTab = tbPageManageTank
                        If MsgBox("Do you want to save the new Pipe?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            pPipe.Remove(nTankID.ToString + "|" + nCompartmentNumber.ToString + "|" + nPipeID.ToString)
                            ResetPipeID()
                            bolAddPipe = False
                            tbCntrlRegistration.SelectedTab = tb
                        Else
                            btnPipeSave.PerformClick()
                            retVal = False
                        End If
                    End If
                End If
            End If
            Return retVal
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Function

    Friend Sub SetupTabs()
        Try
            If Not CheckTabs() Then
                Exit Sub
            End If

            Select Case tbCntrlRegistration.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    tbCntrlRegistration.Tag = tbPageOwnerDetail.Name
                    PopulateOwner(nOwnerID, True)
                    ' clearing the facility info in mustercontainer
                    MC.lblFacilityInfo.Text = String.Empty
                    MC.lblFacilityID.Text = String.Empty
                Case tbPageFacilityDetail.Name
                    If nFacilityID = 0 And nOwnerID = 0 Then
                        MsgBox("Please select or add an Owner first")
                        Exit Sub
                    ElseIf nFacilityID = 0 And nOwnerID <> 0 Then
                        If bolAddFacility Then
                            tbCntrlRegistration.Tag = tbPageFacilityDetail.Name
                            PopulateFacility(nFacilityID)
                        Else
                            ' #3007
                            'If lblFacilityIDValue.Text <> String.Empty Then
                            '    If IsNumeric(lblFacilityIDValue.Text) Then
                            '        nFacilityID = lblFacilityIDValue.Text
                            '        PopulateFacility(nFacilityID)
                            '        Exit Select
                            '    End If
                            'End If
                            If Not ugFacilityList.Rows Is Nothing Then
                                If ugFacilityList.Rows.Count > 0 Then
                                    If ugFacilityList.ActiveRow Is Nothing Then
                                        ugFacilityList.ActiveRow = ugFacilityList.Rows(0)
                                    End If
                                    tbCntrlRegistration.Tag = tbPageFacilityDetail.Name
                                    nFacilityID = ugFacilityList.ActiveRow.Cells("FacilityID").Value
                                    PopulateFacility(nFacilityID)
                                Else
                                    If MsgBox("There are no Facilities for Owner" + vbCrLf + _
                                        "Do you want to add Facilities?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                        tbCntrlRegistration.Tag = tbPageFacilityDetail.Name
                                        PopulateFacility(0)
                                    Else
                                        tbCntrlRegistration.SelectedTab = tbPageOwnerDetail
                                    End If
                                End If

                            Else
                                If MsgBox("There are no Facilities for Owner" + vbCrLf + _
                                    "Do you want to add Facilities?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                    tbCntrlRegistration.Tag = tbPageFacilityDetail.Name
                                    PopulateFacility(0)
                                Else
                                    tbCntrlRegistration.SelectedTab = tbPageOwnerDetail
                                End If
                            End If
                        End If
                    ElseIf nFacilityID <> 0 Then
                        tbCntrlRegistration.Tag = tbPageFacilityDetail.Name
                        PopulateFacility(nFacilityID)
                    End If

                Case tbPageManageTank.Name
                    If nFacilityID > 0 Then
                        If tbCntrlTank.Visible Then
                            tbCntrlRegistration.Tag = tbCntrlTank.Name
                            PopulateTank(nTankID)
                        ElseIf tbCntrlPipe.Visible Then
                            tbCntrlRegistration.Tag = tbCntrlPipe.Name
                            PopulatePipe(nPipeID)
                        Else
                            tbCntrlRegistration.Tag = tbPageManageTank.Name
                            PopulateTankPipeGrid(nFacilityID, True)
                        End If
                    Else
                        Dim tp As TabPage = tbCntrlRegistration.SelectedTab
                        tbCntrlRegistration.SelectedTab = tbPageFacilityDetail
                        'removed on 10/17/2007 by hcao, fixed the endless calling stack of SetupTabs for owner with no facility, no tank.
                        '   tbCntrlRegistration.SelectedTab = tp
                    End If
                Case tbPageSummary.Name
                    Me.Text = "Registration - Owner Summary (" & txtOwnerName.Text & ")"
                    'tbCntrlRegistration.Tag = tbPageSummary.Name
                    UIUtilsGen.PopulateOwnerSummary(pOwn, Me)
            End Select

            MC.AppSemaphores.ActivateAuxControls(Me.Text)

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub tbCntrlRegistration_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCntrlRegistration.SelectedIndexChanged
        If bolLoading Then
            Exit Sub
        End If
        SetupTabs()
    End Sub
    Private Sub tbCntrlRegistration_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlRegistration.Click
        Try
            Select Case tbCntrlRegistration.SelectedTab.Name
                Case tbPageManageTank.Name
                    If nFacilityID <= 0 Then
                        tbCntrlRegistration.SelectedTab = tbPageFacilityDetail
                        Exit Sub
                    Else
                        If Not CheckTabs() Then
                            Exit Sub
                        End If
                        ShowHideTankPipeScreen(False, False)
                        PopulateTankPipeGrid(nFacilityID, True)
                        'If Not dgPipesAndTanks.ActiveRow Is Nothing Then
                        '    If dgPipesAndTanks.ActiveRow.Band.Index > 0 Then
                        '        dgPipesAndTanks2.Rows(dgPipesAndTanks.ActiveRow.ParentRow.Index).ChildBands(0).Rows(dgPipesAndTanks.ActiveRow.Index).Activate()
                        '    End If
                        'End If
                        Me.btnAddTank2.Focus()
                    End If
            End Select
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub LicenseeCompanyDetails(ByVal Licensee_id As Integer, ByVal company_id As Integer, ByVal Licensee_name As String, ByVal company_name As String) Handles oCompanySearch.LicenseeCompanyDetails
        Try
            If strFromCompanySearch = "TANK" Then
                txtLicensee.Text = Licensee_name
                txtTankCompany.Text = company_name
                pTank.ContractorID = company_id
                pTank.LicenseedID = Licensee_id
            ElseIf strFromCompanySearch = "FACILITY" Then
                txtFacilityLicensee.Text = Licensee_name
                txtFacilityCompany.Text = company_name
                pOwn.Facilities.ContractorID = company_id
                pOwn.Facilities.LicenseeID = Licensee_id
            Else
                txtPipeLicensee.Text = Licensee_name
                txtPipeCompanyName.Text = company_name
                pPipe.ContractorID = company_id
                pPipe.LicenseeID = Licensee_id
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub lnkLblNextTank_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextTank.LinkClicked
        Try
            If tbCntrlTank.Visible Then
                If Not bolAddTank Then
                    PopulateTank(GetPrevNextTank(nTankID, True))
                End If
            ElseIf tbCntrlPipe.Visible Then
                If Not bolAddPipe Then
                    PopulatePipe(GetPrevNextPipe(nPipeID, True))
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub lnkLblPrevTank_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevTank.LinkClicked
        Try
            If tbCntrlTank.Visible Then
                If Not bolAddTank Then
                    PopulateTank(GetPrevNextTank(nTankID, False))
                End If
            ElseIf tbCntrlPipe.Visible Then
                If Not bolAddPipe Then
                    PopulatePipe(GetPrevNextPipe(nPipeID, False))
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub AddressForm_NewAddressID(ByVal MyAddressID As Integer) Handles AddressForm.NewAddressID
        Try
            Select Case tbCntrlRegistration.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    txtOwnerAddress.Tag = MyAddressID
                    pOwn.AddressId = MyAddressID
                Case tbPageFacilityDetail.Name
                    txtFacilityAddress.Tag = MyAddressID
                    pOwn.Facilities.AddressID = MyAddressID
            End Select
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region

#Region "Owner Tab"
    Friend Sub PopulateOwner(ByVal ownID As Integer, ByVal resetOtherIDs As Boolean)
        Try
            If resetOtherIDs Then
                ResetOwnerID()
            End If
            nOwnerID = ownID

            ' retrieve owner and populate the ui
            UIUtilsGen.PopulateOwnerInfo(ownID, pOwn, Me)

            If pOwn.ID <= 0 Then
                nOwnerID = pOwn.ID
                Me.Text = "Registration - Owner Detail - (New)"

                Me.btnOwnerComment.Enabled = False
                Me.btnOwnerFlag.Enabled = False
                Me.btnTransferOwnership.Enabled = False
                Me.LinkLblCAPSignup.Enabled = False
                CommentsMaintenance(, , True, True)
            Else
                Me.Text = "Registration" & " - Owner Detail - " & lblOwnerIDValue.Text & " (" & txtOwnerName.Text & ")"

                Me.btnOwnerComment.Enabled = True
                Me.btnOwnerFlag.Enabled = True
                Me.LinkLblCAPSignup.Enabled = True

                If Me.ugFacilityList.Rows.Count > 0 Then
                    Me.btnDeleteOwner.Enabled = False
                    Me.btnTransferOwnership.Enabled = True
                Else
                    Me.btnDeleteOwner.Enabled = True
                    Me.btnTransferOwnership.Enabled = False
                End If
                ' check for pending registration activity
                CheckForRegistrationActivity()
                CommentsMaintenance(, , True)
            End If

            If tbCtrlOwner.SelectedTab.Name = tbPrevFacs.Name Then
                PopulatePreviouslyOwnedFacilities(ownID)
            ElseIf tbCtrlOwner.SelectedTab.Name = tbPageOwnerContactList.Name Then
                LoadContacts(ugOwnerContacts, ownID, UIUtilsGen.EntityTypes.Owner)
            End If

            SetOwnerSaveCancel(pOwn.IsDirty)

            If Me.ugFacilityList.Rows.Count > 0 Then
                Me.btnDeleteOwner.Enabled = False
                Me.btnTransferOwnership.Enabled = True
            Else
                Me.btnDeleteOwner.Enabled = True
                Me.btnTransferOwnership.Enabled = False
            End If

            MC.FlagsChanged(ownID, UIUtilsGen.EntityTypes.Owner, "Registration", Me.Text)
            lblOwnerIDValue.Focus()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub SetOwnerSaveCancel(ByVal bolstate As Boolean)
        btnSaveOwner.Enabled = bolstate
        btnOwnerCancel.Enabled = bolstate
        btnDeleteOwner.Enabled = False
        'If nOwnerID > 0 Then
        '    If Not ugFacilityList.Rows Is Nothing Then
        '        If ugFacilityList.Rows.Count = 0 Then
        '            btnDeleteOwner.Enabled = True
        '        End If
        '    End If
        'End If
    End Sub
    Private Sub SetPersonaSaveCancel(ByVal bolstate As Boolean)
        btnOwnerNameOK.Enabled = bolstate
        btnOwnerNameCancel.Enabled = bolstate
    End Sub
    Friend Sub NewOwner()
        'pOwn.Retrieve(0)
    End Sub
    Friend Sub SetupAddOwner()
        Try
            ClearOwnerForm()
            bolAddOwner = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ClearOwnerForm()
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            UIUtilsGen.ClearFields(pnlOwnerDetail)
            rdOwnerOrg.Tag = False
            rdOwnerPerson.Tag = True
            chkCAPParticipant.Tag = True 'Person Mode is the default for Owner Name
            cmbOwnerType.SelectedIndex = -1
            If cmbOwnerType.SelectedIndex <> -1 Then
                cmbOwnerType.SelectedIndex = -1
            End If
            txtOwnerAddress.Tag = 0
            txtOwnerName.Tag = 0
            lblNoOfFacilitiesValue.Text = "0"
        Catch ex As Exception
            Throw ex
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Friend Sub CallTransferOwnership()
        Try
            If Not ugFacilityList.Rows.Count > 0 Then
                MsgBox("This Owner (" + nOwnerID.ToString + ") has no Facilities to Transfer.")
                Exit Sub
            End If

            Dim ownTransfer As New TransferOwnership(nOwnerID, pOwn)
            'ownTransfer.CurrentOwnerInformation()
            ownTransfer.MdiParent = MC
            ownTransfer.BringToFront()
            ownTransfer.Show()
            'ownTransfer.cmbOwnerRight.SelectedIndex = -1
            'If ownTransfer.cmbOwnerRight.SelectedIndex <> -1 Then
            '    ownTransfer.cmbOwnerRight.SelectedIndex = -1
            'End If

            'If Not ugFacilityList.Rows.Count > 0 Then
            '    MsgBox("This Owner (" + nOwnerID.ToString + ") has no Facilities to Transfer.")
            '    Exit Sub
            'End If
            'Dim ownTransfer As New TransferOwnership(Me, Me.MdiParent, , pOwn)
            'AddHandler ownTransfer.Closing, AddressOf frmTrasferOwnershipClosing
            'ownTransfer.CurrentOwnerInformation()
            'ownTransfer.MdiParent = MC
            'ownTransfer.BringToFront()
            'ownTransfer.Show()
            'ownTransfer.cmbOwner2.SelectedIndex = -1
            'If ownTransfer.cmbOwner2.SelectedIndex <> -1 Then
            '    ownTransfer.cmbOwner2.SelectedIndex = -1
            'End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load currentOwner: " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Friend Sub PopulateOwnerFacs(ByVal ownID As Integer)
        Try
            If nOwnerID = pOwn.ID Then
                UIUtilsGen.PopulateOwnerFacilities(pOwn, Me, ownID)
            Else
                MsgBox("Incorrect Owner (" + ownID.ToString + ")supplied. Form showing Owner (" + nOwnerID.ToString + ")")
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Friend Sub PopulatePreviouslyOwnedFacilities(ByVal ownID As Integer)
        Try
            If tbCtrlOwner.SelectedTab.Name <> tbPrevFacs.Name Then
                tbCtrlOwner.SelectedTab = tbPrevFacs
                Exit Sub
            End If
            ugPrevOwnedFacs.DataSource = pOwn.PreviousFacilities()
            If Not ugPrevOwnedFacs.Rows Is Nothing Then
                If Not ugPrevOwnedFacs.Rows.Count > 0 Then
                    MsgBox("No Previous Facilities exists for Owner: " + ownID.ToString)
                    If tbCtrlOwner.Tag Is Nothing Then
                        tbCtrlOwner.SelectedTab = tbPageOwnerFacilities
                    ElseIf tbCtrlOwner.Tag = tbPageOwnerContactList.Name Then
                        tbCtrlOwner.SelectedTab = tbPageOwnerContactList
                    Else
                        tbCtrlOwner.SelectedTab = tbPageOwnerFacilities
                    End If
                Else
                    lblPreviousFacilitiesCountValue.Text = ugPrevOwnedFacs.Rows.Count.ToString
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ugPrevOwnedFacs_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugPrevOwnedFacs.InitializeLayout
        Try
            e.Layout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
            e.Layout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            e.Layout.Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
            e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            e.Layout.Bands(0).Columns("From").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("To").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(0).Columns("FacilityID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            e.Layout.Bands(0).Columns("Facility Name").Width = 200
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub ugPrevOwnedFacs_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugPrevOwnedFacs.DoubleClick
        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Try
            If Not ugPrevOwnedFacs.ActiveRow Is Nothing Then
                If Not ugPrevOwnedFacs.ActiveRow.Cells("FacilityID").Value Is DBNull.Value Then
                    MC.txtOwnerQSKeyword.Text = ugPrevOwnedFacs.ActiveRow.Cells("FacilityID").Value
                    MC.cmbSearchModule.SelectedIndex = MC.cmbSearchModule.FindString(UIUtilsGen.ModuleID.Registration.ToString)
                    MC.cmbQuickSearchFilter.SelectedIndex = MC.cmbQuickSearchFilter.FindString("Facility ID")
                    MC.btnQuickOwnerSearch.PerformClick()
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub tbCtrlOwner_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlOwner.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            Select Case tbCtrlOwner.SelectedTab.Name
                Case tbPrevFacs.Name
                    PopulatePreviouslyOwnedFacilities(nOwnerID)
                Case tbPageOwnerContactList.Name
                    LoadContacts(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
                Case tbPageOwnerDocuments.Name
                    UCOwnerDocuments.LoadDocumentsGrid(nOwnerID, 9, 612)
                Case tbPageOwnerFacilities.Name
                    PopulateOwnerFacs(nOwnerID)
            End Select
            tbCtrlOwner.Tag = tbCtrlOwner.SelectedTab.Name
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub txtOwnerAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.TextChanged
        If bolLoading Then Exit Sub
        If txtOwnerAddress.Tag > 0 Then
            pOwn.AddressId = Integer.Parse(Trim(txtOwnerAddress.Tag))
        Else
            pOwn.AddressId = 0
        End If
    End Sub
    Private Sub txtOwnerAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.DoubleClick
        Try
            AddressForm = New Address(UIUtilsGen.EntityTypes.Owner, pOwn.Addresses, "Owner", , UIUtilsGen.ModuleID.Registration)
            AddressForm.ShowCounty = False
            AddressForm.ShowFIPS = False
            AddressForm.ShowDialog()
            ' update address text
            pOwn.Addresses.Retrieve(pOwn.AddressId)
            txtOwnerAddress.Text = UIUtilsGen.FormatAddress(pOwn.Addresses)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerAddress_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOwnerAddress.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                txtOwnerAddress_DoubleClick(sender, New System.EventArgs)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerAddress_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.Enter
        Try
            If txtOwnerAddress.Text = String.Empty Then
                txtOwnerAddress_DoubleClick(sender, e)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmbOwnerType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        pOwn.OwnerType = UIUtilsGen.GetComboBoxValueInt(cmbOwnerType)
    End Sub
    Private Sub txtOwnerEmail_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerEmail.TextChanged
        If bolLoading Then Exit Sub
        pOwn.EmailAddress = Me.txtOwnerEmail.Text
    End Sub
    Private Sub chkOwnerAgencyInterest_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOwnerAgencyInterest.CheckedChanged
        If bolLoading Then Exit Sub
        pOwn.EnsiteAgencyInterestID = chkOwnerAgencyInterest.Checked
    End Sub
    Private Sub txtOwnerAIID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerAIID.TextChanged
        If bolLoading Then Exit Sub
        Dim nAIID As Integer = 0
        Try
            If txtOwnerAIID.Text <> String.Empty Then
                If IsNumeric(txtOwnerAIID.Text) Then
                    nAIID = CType(txtOwnerAIID.Text, Integer)
                Else
                    MsgBox("Please enter valid numeric value")
                    txtOwnerAIID.Text = String.Empty
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox("Invalid Entry")
            If pOwn.EnsiteOrganizationID <> 0 And pOwn.EnsitePersonID = 0 Then
                txtOwnerAIID.Text = pOwn.EnsiteOrganizationID.ToString
            ElseIf pOwn.EnsiteOrganizationID = 0 And pOwn.EnsitePersonID <> 0 Then
                txtOwnerAIID.Text = pOwn.EnsitePersonID.ToString
            Else
                txtOwnerAIID.Text = String.Empty
            End If
        End Try
        If pOwn.OrganizationID > 0 And pOwn.PersonID = 0 Then
            pOwn.EnsiteOrganizationID = nAIID
            pOwn.EnsitePersonID = 0
        ElseIf pOwn.OrganizationID = 0 And pOwn.PersonID > 0 Then
            pOwn.EnsiteOrganizationID = 0
            pOwn.EnsitePersonID = nAIID
        End If
    End Sub
    Private Sub mskTxtOwnerPhone_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtOwnerPhone.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pOwn.PhoneNumberOne, mskTxtOwnerPhone.FormattedText.Trim.ToString)
    End Sub
    Private Sub mskTxtOwnerPhone2_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtOwnerPhone2.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pOwn.PhoneNumberTwo, mskTxtOwnerPhone2.FormattedText.Trim.ToString)
    End Sub
    Private Sub mskTxtOwnerFax_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtOwnerFax.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pOwn.Fax, mskTxtOwnerFax.FormattedText.Trim.ToString)
    End Sub

    ' Owner Name
    Private Sub SwapOrgPersonDisplay()
        Try
            UIUtilsGen.SwapOrgPersonDisplay(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function CheckObjectState(ByRef obj As MUSTER.BusinessLogic.pPersona, Optional ByVal bolColDirty As Boolean = False) As Boolean ', Optional ByVal sender As Object = Nothing, Optional ByVal e As System.ComponentModel.CancelEventArgs = Nothing)
        CheckObjectState = False
        Try
            If Not obj Is Nothing Then
                If obj.IsDirty Then
                    Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                    If Results = MsgBoxResult.Yes Then
                        Dim success As Boolean = False
                        If obj.PersonId > 0 Or obj.OrgID Then

                        End If
                        returnVal = String.Empty
                        success = obj.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Function
                        End If

                        If Not success Then
                            Exit Function
                        End If
                        UIUtilsGen.SetOwnerName(Me)
                    ElseIf Results = MsgBoxResult.No Then
                        Dim evt As System.EventArgs
                        If obj.GetType.ToString.ToLower = "muster.businesslogic.ppersona" Then
                            Me.btnOwnerNameCancel.PerformClick()
                        End If
                    ElseIf Results = MsgBoxResult.Cancel Then
                        Dim e As System.ComponentModel.CancelEventArgs
                        If Not obj.GetType.ToString.ToLower = "muster.businesslogic.ppersona" Then
                            e.Cancel = True
                        Else
                            CheckObjectState = True
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub txtOwnerName_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerName.DoubleClick
        Try
            If Me.txtOwnerName.Text <> String.Empty Then
                UIUtilsGen.setupOwnername(Me, pOwn)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerName.Enter
        Try
            UIUtilsGen.setupOwnername(Me, pOwn)
            If pOwn.PersonID = 0 And pOwn.OrganizationID <> 0 Then
                txtOwnerOrgName.Focus()
            Else
                cmbOwnerNameTitle.Focus()
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerName.TextChanged
        If bolLoading Then Exit Sub
        Try
            If pOwn.BPersona.OrgID <> 0 Then
                pOwn.OrganizationID = pOwn.BPersona.OrgID
                If pOwn.PersonID > 0 Then
                    pOwn.PersonID = 0
                End If
            Else
                pOwn.PersonID = pOwn.BPersona.PersonId
                If pOwn.OrganizationID > 0 Then
                    pOwn.OrganizationID = 0
                End If
            End If
            SetOwnerSaveCancel(True)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub rdOwnerPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdOwnerPerson.Click
        Try
            UIUtilsGen.rdOwnerPersonClick(Me, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub rdOwnerOrg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdOwnerOrg.Click
        Try
            UIUtilsGen.rdOwnerOrgClick(Me, pOwn)
            If rdOwnerOrg.Checked And pOwn.BPersona.Org_Entity_Code = 0 Then
                cmbOwnerOrgEntityCode.SelectedValue = 539  ' Default "UnderGround Storage Tank Owner"
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbOwnerNameTitle_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerNameTitle.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(cmbOwnerNameTitle, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerFirstName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerFirstName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerFirstName, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerMiddleName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerMiddleName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerMiddleName, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerLastName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerLastName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerLastName, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbOwnerNameSuffix_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerNameSuffix.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(cmbOwnerNameSuffix, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerOrgName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerOrgName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerOrgName, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbOwnerOrgEntityCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerOrgEntityCode.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(cmbOwnerOrgEntityCode, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub pnlPersonOrganization_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlPersonOrganization.LostFocus
        Try
            cmbOwnerNameTitle.Focus()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerNameOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerNameOK.Click
        Try
            Dim success As Boolean = False
            If pOwn.BPersona.PersonId > 0 Or pOwn.BPersona.OrgID > 0 Then
                pOwn.BPersona.ModifiedBy = MC.AppUser.ID
            Else
                pOwn.BPersona.CreatedBy = MC.AppUser.ID
            End If
            returnVal = String.Empty
            success = pOwn.BPersona.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            If success Then
                UIUtilsGen.SetOwnerName(Me)
                txtOwnerName.Tag = Nothing
                pnlOwnerName.Hide()
                lblOwnerAddress.Focus()
            Else
                'bolValidateSuccess = True
                'bolDisplayErrmessage = True
                Exit Sub
            End If
            txtOwnerAddress.Focus()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerNameClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerNameClose.Click
        Try
            If Not CheckObjectState(pOwn.BPersona) Then 'And bolValidateSuccess Then
                pnlOwnerName.Hide()
            End If
            'bolValidateSuccess = True
            'bolDisplayErrmessage = True
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerNameClose_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerNameClose.LostFocus
        If Me.rdOwnerOrg.Checked = True Then
            Me.txtOwnerOrgName.Focus()
        Else
            Me.cmbOwnerNameTitle.Focus()
        End If
    End Sub
    Private Sub btnOwnerNameCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerNameCancel.Click
        Try
            If Not pOwn.BPersona Is Nothing Then
                pOwn.BPersona.Reset()
                UIUtilsGen.ClearPersona(Me)
                If pOwn.BPersona.PersonId <> 0 Or pOwn.BPersona.OrgID <> 0 Then
                    UIUtilsGen.ResetOwnerName(Me, pOwn)
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnSaveOwner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveOwner.Click
        'Dim bolAddFac As Boolean = False
        Try
            Dim success As Boolean = False
            pOwn.ModifiedBy = MC.AppUser.ID
            returnVal = String.Empty
            success = pOwn.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If success Then
                bolAddOwner = False
                MsgBox("Owner Saved Successfully")
                If nOwnerID <= 0 Then
                    nOwnerID = pOwn.ID
                    PutRegistrationActivity(nOwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ActivityTypes.AddOwner)
                    If MsgBox("Do you want to add Facilities?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        bolAddFacility = True
                    End If
                End If
                'nOwnerID = pOwn.ID
                PopulateOwner(pOwn.ID, True)

                If bolAddFacility Then
                    nFacilityID = 0
                    tbCntrlRegistration.SelectedTab = tbPageFacilityDetail
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerCancel.Click
        Try
            If Not pOwn Is Nothing Then
                pOwn.Reset()
                ClearOwnerForm()
                If nOwnerID > 0 Then
                    PopulateOwner(nOwnerID, False)
                Else
                    pOwn.BPersona.Clear()
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnDeleteOwner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteOwner.Click
        Try
            If MessageBox.Show("Are you sure you wish to DELETE the Owner?", "DELETE OWNER", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Exit Sub
            End If

            If nOwnerID <= 0 Then
                pOwn = New MUSTER.BusinessLogic.pOwner
            Else
                ' if owner has any fees / fees history, cannot delete owner
                Dim ds As DataSet
                ds = pOwn.RunSQLQuery("SELECT COUNT(OWNER_ID) AS OWNERCOUNT FROM tblFEES_INVOICES WHERE OWNER_ID = " + nOwnerID.ToString)
                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        If ds.Tables(0).Rows(0)("OWNERCOUNT") > 0 Then
                            MsgBox("Owner has Fees associated. Cannot delete")
                            Exit Sub
                        End If
                    End If
                End If

                If pOwn.IsDirty Then
                    If MessageBox.Show("There are unsaved changes.Do you want to save the changes before delete? ", "Caption", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                        pOwn.Reset()
                    End If
                End If
                pOwn.Deleted = True
                pOwn.ModifiedBy = MC.AppUser.ID
                Dim success As Boolean = False
                returnVal = String.Empty
                success = pOwn.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                If success Then
                    DeleteRegistrationActivities(nOwnerID)
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnTransferOwnership_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTransferOwnership.Click
        Try
            CallTransferOwnership()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub ugFacilityList_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugFacilityList.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            nFacilityID = ugFacilityList.ActiveRow.Cells("FacilityID").Value
            tbCntrlRegistration.SelectedTab = tbPageFacilityDetail
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub LinkLblCAPSignup_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLblCAPSignup.LinkClicked
        Dim frmCAP As CAPSignUp
        Try
            frmCAP = New CAPSignUp(pOwn)
            frmCAP.nOwnerID = nOwnerID
            If lblFacilityIDValue.Text <> String.Empty Then
                If IsNumeric(lblFacilityIDValue.Text) Then frmCAP.nFacilityID = lblFacilityIDValue.Text
            End If
            frmCAP.txtOwner.Text = Me.txtOwnerName.Text + " (" + Me.lblOwnerIDValue.Text + ")"
            'frmCAP.txtOwner.Tag = Me.txtOwnerName.Text
            Me.Tag = "0"
            frmCAP.CallingForm = Me
            frmCAP.ShowDialog()
            If Me.Tag = "1" Then
                lblCAPParticipationLevel.Text = pOwn.CAPParticipationLevel
                ' clearing the collection
                pOwn.OwnerInfo.facilityCollection = New MUSTER.Info.FacilityCollection
                pOwn.Facilities = New MUSTER.BusinessLogic.pFacility
                pTank = New MUSTER.BusinessLogic.pTank
                pPipe = New MUSTER.BusinessLogic.pPipe
                MC.FlagsChanged(nOwnerID, UIUtilsGen.EntityTypes.Owner, "Registration", Me.Text)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region

#Region "Facility Tab"
    Friend Sub PopulateFacility(ByVal facID As Integer)
        Try
            ResetFacilityID()
            nFacilityID = facID

            MC.lblFacilityInfo.Text = String.Empty
            MC.lblFacilityID.Text = String.Empty
            If nFacilityID = 0 Then
                pOwn.Facilities.Retrieve(pOwn.OwnerInfo, 0, , "FACILITY", False, True)
                ' #2802
                If lblCAPParticipationLevel.Text.StartsWith("FULL") Or lblCAPParticipationLevel.Text.StartsWith("PARTIAL") Then
                    pOwn.Facilities.CAPCandidate = True
                End If
            End If
            UIUtilsGen.PopulateFacilityInfo(Me, pOwn.OwnerInfo, pOwn.Facilities, facID)
            txtFacilityNameForEnsite.Text = pOwn.Facilities.NameForEnsite
            If pOwn.Facilities.ID <= 0 Then
                nFacilityID = pOwn.Facilities.ID
                Me.Text = "Registration - Facility Detail - (New)"
                btnDeleteFacility.Enabled = False
                dgPipesAndTanks.DataSource = Nothing
                dgPipesAndTanks2.DataSource = Nothing
                CommentsMaintenance(, , True, True)
            Else
                Me.Text = "Registration" & " - Facility Detail - " & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"

                PopulateTankPipeGrid(facID, False)

                If tabCtrlFacilityTankPipe.SelectedTab.Name = tpPreviouslyOwnedOwners.Name Then
                    PopulatePreviouslyOwnedOwners(facID)
                ElseIf tabCtrlFacilityTankPipe.SelectedTab.Name = tbPageFacilityContactList.Name Then
                    LoadContacts(ugFacilityContacts, facID, UIUtilsGen.EntityTypes.Facility)
                ElseIf tabCtrlFacilityTankPipe.SelectedTab.Name = tbPageFacilityDocuments.Name Then
                    UCFacilityDocuments.LoadDocumentsGrid(nFacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Registration)
                End If
                btnDeleteFacility.Enabled = True
                ' check for pending registration activity
                CheckForRegistrationActivity()
                CommentsMaintenance(, , True)
            End If
            FillFacilityLicenseeDetails()
            SetFacilitySaveCancel(pOwn.Facilities.IsDirty)
            ShowHideTankPipeScreen(False, False)
            MC.FlagsChanged(facID, UIUtilsGen.EntityTypes.Facility, "Registration", Me.Text)
            txtFacilityName.Focus()

            'Added by Hua Cao 10/16/2008 Issue #: [ UST-3204] Summary: Need a date field labeled " TOS Assessment Date:" added to several modules
            ' Retreive info from tblReg_AssessDate
            If Convert.ToInt32(lblFacilityIDValue.Text) > 0 Then
                Dim sqlStr As String
                Dim dtReturn As DataTable
                Me.dtPickAssess.Enabled = True
                sqlStr = "tblReg_AssessDate where FacilityId = " + lblFacilityIDValue.Text
                dtReturn = pTank.GetDataTable(sqlStr)
                If Not dtReturn Is Nothing Then
                    If dtReturn.Rows.Count > 0 Then
                        If Not dtReturn.Rows(0).Item("AssessDate") Is System.DBNull.Value Then
                            If btnFacilitySave.Enabled = False Then
                                Me.dtPickAssess.Value = dtReturn.Rows(0).Item("AssessDate")
                                btnFacilitySave.Enabled = False
                                btnFacilityCancel.Enabled = False
                            Else
                                Me.dtPickAssess.Value = dtReturn.Rows(0).Item("AssessDate")
                            End If
                        Else
                            Me.dtPickAssess.Checked = False
                            UIUtilsGen.SetDatePickerValue(dtPickAssess, dtNullDate)
                        End If
                    End If
                    dtReturn.Clear()
                Else
                    Me.dtPickAssess.Checked = False
                    UIUtilsGen.SetDatePickerValue(dtPickAssess, dtNullDate)
                End If
            Else
                Me.dtPickAssess.Checked = False
                UIUtilsGen.SetDatePickerValue(dtPickAssess, dtNullDate)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub SetFacilitySaveCancel(ByVal bolState As Boolean)
        btnFacilitySave.Enabled = bolState
        btnFacilityCancel.Enabled = bolState
    End Sub
    Private Function GetPrevNextFacility(ByVal facID As Integer, ByVal getNext As Boolean) As Integer
        Try
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim retVal As Integer = facID
            Dim sl As New SortedList

            For Each ugRow In ugFacilityList.Rows
                sl.Add(ugRow.Cells("FacilityID").Value, ugRow.Cells("FacilityID").Value)
            Next
            Return GetPrevNext(sl, getNext, facID)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Function
    Private Sub Check_If_Datum_Enable(Optional ByVal sender As Object = Nothing, Optional ByVal e As System.ComponentModel.CancelEventArgs = Nothing)
        Try
            UIUtilsGen.Check_If_Datum_Enable(Me)
            ' UST bug 52
            ' setting default values to datum, method and type
            ' by setting selected value, it changed the drop down's selected index and sets the value in the object
            If cmbFacilityDatum.Enabled Then
                If pOwn.Facilities.Datum = 0 Then
                    cmbFacilityDatum.SelectedValue = 581
                    'pOwn.Facilities.Datum = 581
                End If
            End If
            If cmbFacilityMethod.Enabled Then
                If pOwn.Facilities.Method = 0 Then
                    'pOwn.Facilities.Method = 583
                    cmbFacilityMethod.SelectedValue = 583
                End If
            End If
            If cmbFacilityLocationType.Enabled Then
                If pOwn.Facilities.LocationType = 0 Then
                    'pOwn.Facilities.LocationType = 588
                    cmbFacilityLocationType.SelectedValue = 588
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FillFacilityLicenseeDetails()
        Try
            If Date.Compare(pOwn.Facilities.UpcomingInstallationDate, dtNullDate) = 0 Then
                txtFacilityLicensee.Enabled = False
                txtFacilityCompany.Enabled = False
                lblFacilityLicenseeSearch.Enabled = False

                pOwn.Facilities.LicenseeID = 0
                txtFacilityLicensee.Text = String.Empty

                pOwn.Facilities.ContractorID = 0
                txtFacilityCompany.Text = String.Empty
            Else
                txtFacilityLicensee.Enabled = True
                txtFacilityCompany.Enabled = True
                lblFacilityLicenseeSearch.Enabled = True

                pLicensee.Retrieve(pOwn.Facilities.LicenseeID)
                txtFacilityLicensee.Text = pLicensee.Licensee_name

                pCompany.Retrieve(pOwn.Facilities.ContractorID)
                txtFacilityCompany.Text = pCompany.COMPANY_NAME
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Friend Sub PopulateTankPipeGrid(ByVal facID As Integer, ByVal setFormText As Boolean)
        Try
            dgPipesAndTanks.DataSource = pOwn.Facilities.TankPipeDataset(facID)
            ExpandAll(False, dgPipesAndTanks, btnExpand)

            lblTotalNoOfTanksValue.Text = dgPipesAndTanks.Rows.Count.ToString
            lblTankCountVal.Text = " of " + lblTotalNoOfTanksValue.Text
            lblTankCountVal.Left = lblTankIDValue.Left + lblTankIDValue.Width

            PopulateTankPipeGrid2()

            ugTankRow = Nothing
            ugPipeRow = Nothing

            Dim allTanksPipesPOU As Boolean = True
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In dgPipesAndTanks.Rows
                If ugRow.Cells("STATUS").Text <> "Permanently Out of Use" Then
                    allTanksPipesPOU = False
                    Exit For
                End If
                For Each ugRowChild As Infragistics.Win.UltraWinGrid.UltraGridRow In ugRow.ChildBands(0).Rows
                    If ugRowChild.Cells("PIPE STATUS").Text <> "Permanently Out of Use" Then
                        allTanksPipesPOU = False
                        Exit For
                    End If
                Next
                If Not allTanksPipesPOU Then Exit For
            Next
            If allTanksPipesPOU Then
                If chkCAPCandidate.Checked Then
                    MsgBox("Facility cannot be CAP Candidate as all Tanks / Pipes are POU")
                End If
                'Commented out by Hua Cao on Nov. 7, 2007. 
                'Alternative solution: enable the chkCAPCandidate
                chkCAPCandidate.Checked = False
                chkCAPCandidate.Enabled = False
                lblCAPStatusValue.Text = String.Empty
                lblCAPStatusValue.BackColor = System.Drawing.SystemColors.Control
                pOwn.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)
            Else
                chkCAPCandidate.Enabled = True
                UIUtilsGen.PopulateFacilityCapInfo(Me, pOwn.Facilities)
            End If

            If setFormText Then
                Me.Text = "Registration - Manage Tank/Pipe Summary - " + lblFacilityIDValue.Text + " (" + txtFacilityName.Text + ")"
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateTankPipeGrid2()
        Try
            dgPipesAndTanks2.DataSource = dgPipesAndTanks.DataSource
            ExpandAll(True, dgPipesAndTanks2, btnExpandTP2)

            lblTotalNoOfTanksValue2.Text = dgPipesAndTanks2.Rows.Count.ToString
            lblTankCountVal2.Text = lblTankCountVal.Text
            lblTankCountVal2.Left = lblTankIDValue2.Left + lblTankIDValue2.Width
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub dgPipesAndTanks_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles dgPipesAndTanks.InitializeLayout
        Try
            e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            e.Layout.Bands(0).Columns("FACILITY_ID").Hidden = True
            e.Layout.Bands(0).Columns("TANK ID").Hidden = True
            e.Layout.Bands(0).Columns("COMPARTMENT").Hidden = True
            e.Layout.Bands(0).Columns("POSITION").Hidden = True

            e.Layout.Bands(0).Columns("INSTALLED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("LAST USED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.AutoFitColumns = True

            If e.Layout.Bands.Count > 1 Then
                e.Layout.Bands(1).Columns("FACILITY_ID").Hidden = True
                e.Layout.Bands(1).Columns("TANK ID").Hidden = True
                e.Layout.Bands(1).Columns("COMPARTMENT NUMBER").Hidden = False
                e.Layout.Bands(1).Columns("PIPE ID").Hidden = True
                e.Layout.Bands(1).Columns("FILLER").Hidden = True
                e.Layout.Bands(1).Columns("POSITION").Hidden = True
                e.Layout.Bands(1).Columns("FILLER2").Hidden = True

                e.Layout.Bands(1).Columns("INSTALL DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(1).Columns("LAST USED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dgPipesAndTanks2_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles dgPipesAndTanks2.InitializeLayout
        Try
            e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            e.Layout.Bands(0).Columns("FACILITY_ID").Hidden = True
            e.Layout.Bands(0).Columns("TANK ID").Hidden = True
            e.Layout.Bands(0).Columns("COMPARTMENT").Hidden = True
            e.Layout.Bands(0).Columns("POSITION").Hidden = True

            e.Layout.Bands(0).Columns("INSTALLED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("LAST USED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.AutoFitColumns = True

            If e.Layout.Bands.Count > 1 Then
                e.Layout.Bands(1).ColHeadersVisible = False
                e.Layout.Bands(1).Columns("COMPARTMENT NUMBER").Hidden = False
                e.Layout.Bands(1).Columns("FACILITY_ID").Hidden = True
                e.Layout.Bands(1).Columns("TANK ID").Hidden = True
                e.Layout.Bands(1).Columns("PIPE ID").Hidden = True
                e.Layout.Bands(1).Columns("POSITION").Hidden = True
                e.Layout.Bands(1).Columns("FILLER2").Hidden = True

                e.Layout.Bands(1).Columns("INSTALL DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                e.Layout.Bands(1).Columns("LAST USED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dgPipesAndTanksDoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim ug As Infragistics.Win.UltraWinGrid.UltraGrid = CType(sender, Infragistics.Win.UltraWinGrid.UltraGrid)
            If ug.ActiveRow.Band.Index = 0 Then
                ugTankRow = ug.ActiveRow
                nTankID = ug.ActiveRow.Cells("TANK ID").Value
                ShowHideTankPipeScreen(True, False)
                tbCntrlRegistration.SelectedTab = tbPageManageTank
            Else
                ugPipeRow = ug.ActiveRow
                nTankID = ug.ActiveRow.Cells("TANK ID").Value
                nCompartmentNumber = ug.ActiveRow.Cells("COMPARTMENT NUMBER").Value
                nPipeID = ug.ActiveRow.Cells("PIPE ID").Value
                ShowHideTankPipeScreen(False, True)
                tbCntrlRegistration.SelectedTab = tbPageManageTank
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub dgPipesAndTanks_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgPipesAndTanks.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            dgPipesAndTanksDoubleClick(sender, e)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dgPipesAndTanks2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgPipesAndTanks2.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            dgPipesAndTanksDoubleClick(sender, e)
            SetupTabs()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Friend Sub PopulatePreviouslyOwnedOwners(ByVal facID As Integer)
        Try
            If tabCtrlFacilityTankPipe.SelectedTab.Name <> tpPreviouslyOwnedOwners.Name Then
                tabCtrlFacilityTankPipe.SelectedTab = tpPreviouslyOwnedOwners
                Exit Sub
            End If
            ugPreviouslyOwnedOwners.DataSource = pOwn.Facilities.PreviousOwners()
            If Not ugPreviouslyOwnedOwners.Rows Is Nothing Then
                If Not ugPreviouslyOwnedOwners.Rows.Count > 0 Then
                    MsgBox("No Previous Owners exists for Facility: " + facID.ToString)
                    If tabCtrlFacilityTankPipe.Tag Is Nothing Then
                        tabCtrlFacilityTankPipe.SelectedTab = tbPageTankPipe
                    ElseIf tabCtrlFacilityTankPipe.Tag = tbPageFacilityContactList.Name Then
                        tabCtrlFacilityTankPipe.SelectedTab = tbPageFacilityContactList
                    Else
                        tabCtrlFacilityTankPipe.SelectedTab = tbPageTankPipe
                    End If
                Else
                    lblNoofOwnersValue.Text = ugPreviouslyOwnedOwners.Rows.Count.ToString
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ugPreviouslyOwnedOwners_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugPreviouslyOwnedOwners.InitializeLayout
        Try
            e.Layout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
            e.Layout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            e.Layout.Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
            e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            e.Layout.Bands(0).Columns("From").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("To").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(0).Columns("To").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
            e.Layout.Bands(0).Columns("Owner Name").Width = 200
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub ugPreviouslyOwnedOwners_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugPreviouslyOwnedOwners.DoubleClick
        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Try
            If Not ugPreviouslyOwnedOwners.ActiveRow Is Nothing Then
                If Not ugPreviouslyOwnedOwners.ActiveRow.Cells("OwnerID").Value Is DBNull.Value Then
                    MC.txtOwnerQSKeyword.Text = ugPreviouslyOwnedOwners.ActiveRow.Cells("OwnerID").Value
                    MC.cmbSearchModule.SelectedIndex = MC.cmbSearchModule.FindString(UIUtilsGen.ModuleID.Registration.ToString)
                    MC.cmbQuickSearchFilter.SelectedIndex = MC.cmbQuickSearchFilter.FindString("Owner ID")
                    MC.btnQuickOwnerSearch.PerformClick()
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub tabCtrlFacilityTankPipe_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCtrlFacilityTankPipe.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            Select Case tabCtrlFacilityTankPipe.SelectedTab.Name
                Case tpPreviouslyOwnedOwners.Name
                    PopulatePreviouslyOwnedOwners(nFacilityID)
                Case tbPageFacilityContactList.Name
                    LoadContacts(ugFacilityContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(nFacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Registration)
                Case tbPageTankPipe.Name
                    PopulateTankPipeGrid(nFacilityID, False)
            End Select
            tabCtrlFacilityTankPipe.Tag = tabCtrlFacilityTankPipe.SelectedTab.Name
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub tabCtrlFacilityTankPipe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCtrlFacilityTankPipe.Click
        If bolLoading Then Exit Sub
        Try
            Select Case tabCtrlFacilityTankPipe.SelectedTab.Name
                Case tbPageFacilityContactList.Name
                    LoadContacts(ugFacilityContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(nFacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Registration)
            End Select
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub txtFacilityName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityName.TextChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.Name = txtFacilityName.Text
            ' TODO - is this required. it is better to handle this when facility is transferred to another owner
            'If Me.lblDateTransfered.Text <> String.Empty Then
            '    Dim dtDateTransferred As Date = CDate(lblDateTransfered.Text)
            '    If Date.Compare(pOwn.Facilities.DateTransferred, dtDateTransferred) <> 0 Then
            '        pOwn.Facilities.DateTransferred = dtDateTransferred 'CType(Trim(lblDateTransfered.Text), Date)
            '    End If
            'End If
            ' ownerid is set when a new instance is created
            'If Me.txtFacilityName.Text <> String.Empty Then
            '    pOwn.Facilities.OwnerID = Integer.Parse(lblOwnerIDValue.Text.Trim)
            'Else
            '    pOwn.Facilities.OwnerID = 0
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityName.Leave
        If bolLoading Then Exit Sub
        Try
            If txtFacilityNameForEnsite.Text = String.Empty Then
                txtFacilityNameForEnsite.Text = txtFacilityName.Text
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityNameForEnsite_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityNameForEnsite.TextChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.NameForEnsite = txtFacilityNameForEnsite.Text
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub mskTxtFacilityPhone_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtFacilityPhone.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pOwn.Facilities.Phone, mskTxtFacilityPhone.FormattedText.ToString)
    End Sub
    Private Sub mskTxtFacilityFax_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtFacilityFax.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pOwn.Facilities.Fax, mskTxtFacilityFax.FormattedText.ToString)
    End Sub
    Private Sub chkSignatureofNF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSignatureofNF.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            ' to be handled on facility save
            'If chkSignatureofNF.Checked Then
            '    If nFacilityID > 0 Then
            '        Dim xCalInfo As MUSTER.Info.CalendarInfo
            '        Dim colFlags As MUSTER.Info.FlagsCollection

            '        colFlags = MC.pFlag.RetrieveFlags(nFacilityID, uiutilsgen.EntityTypes.Facility, , , , , , "Signature Required Letter")
            '        For Each flagInfo As MUSTER.Info.FlagInfo In colFlags.Values
            '            flagInfo.Deleted = True
            '            'To Delete the System Generated SignatureNF DUE TO ME.
            '            xCalInfo = MC.pCalendar.Retrieve(flagInfo.CalendarInfoID)
            '            If xCalInfo.CalendarInfoId > 0 Then
            '                MC.pCalendar.Deleted = True
            '                MC.pCalendar.Save()
            '            End If
            '        Next

            '        If colFlags.Count > 0 Then
            '            MC.pFlag.Flush()
            '        End If
            '    End If
            '    SignatureFlag = False
            'Else
            'End If
            pOwn.Facilities.SignatureOnNF = chkSignatureofNF.Checked
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickFacilityRecvd_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickFacilityRecvd.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickFacilityRecvd)
            pOwn.Facilities.DateReceived = UIUtilsGen.GetDatePickerValue(dtPickFacilityRecvd).Date
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkCAPCandidate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCAPCandidate.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.CAPCandidate = chkCAPCandidate.Checked
            ' will refresh the cap only after facility is saved, else no need
            'If pOwn.Facilities.CAPCandidate And Not pOwn.Facilities.CAPCandidateOriginal Then
            '    pOwn.Facilities.GetCapStatus(pOwn.Facilities.ID)
            '    UIUtilsGen.PopulateFacilityCapInfo(Me, pOwn.Facilities)
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFuelBrand_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFuelBrand.TextChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.FuelBrand = txtFuelBrand.Text
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkUpcomingInstall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUpcomingInstall.Click
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.UpcomingInstallation = chkUpcomingInstall.Checked
            If chkUpcomingInstall.Checked Then
                dtPickUpcomingInstallDateValue.Enabled = True
            Else
                If MsgBox("Do you want to Clear the Date Upcoming Install Date?", MessageBoxButtons.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If

                pOwn.Facilities.UpcomingInstallationDate = dtNullDate

                dtPickUpcomingInstallDateValue.Enabled = False
                UIUtilsGen.SetDatePickerValue(dtPickUpcomingInstallDateValue, dtNullDate)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickUpcomingInstallDateValue_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickUpcomingInstallDateValue.EnabledChanged
        If bolLoading Then Exit Sub
        Try
            If Not dtPickUpcomingInstallDateValue.Enabled Then
                UIUtilsGen.CreateEmptyFormatDatePicker(dtPickUpcomingInstallDateValue)
                pOwn.Facilities.UpcomingInstallationDate = dtNullDate
            End If
            FillFacilityLicenseeDetails()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickUpcomingInstallDateValue_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickUpcomingInstallDateValue.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickUpcomingInstallDateValue)
            pOwn.Facilities.UpcomingInstallationDate = UIUtilsGen.GetDatePickerValue(dtPickUpcomingInstallDateValue).Date
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbFacilityType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFacilityType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.FacilityType = UIUtilsGen.GetComboBoxValueInt(cmbFacilityType)
            txtFacilitySIC.Text = cmbFacilityType.Text
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityLatDegree_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLatDegree.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillSingleObjectValues(pOwn.Facilities.LatitudeDegree, txtFacilityLatDegree.Text.Trim)
            Check_If_Datum_Enable()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityLatMin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLatMin.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillSingleObjectValues(pOwn.Facilities.LatitudeMinutes, txtFacilityLatMin.Text.Trim)
            Check_If_Datum_Enable()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityLatSec_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLatSec.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillDoubleObjectValues(pOwn.Facilities.LatitudeSeconds, txtFacilityLatSec.Text.Trim)
            Check_If_Datum_Enable()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityLongDegree_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLongDegree.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillSingleObjectValues(pOwn.Facilities.LongitudeDegree, txtFacilityLongDegree.Text.Trim)
            Check_If_Datum_Enable()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityLongMin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLongMin.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillSingleObjectValues(pOwn.Facilities.LongitudeMinutes, txtFacilityLongMin.Text.Trim)
            Check_If_Datum_Enable()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityLongSec_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLongSec.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillDoubleObjectValues(pOwn.Facilities.LongitudeSeconds, txtFacilityLongSec.Text.Trim)
            Check_If_Datum_Enable()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbFacilityDatum_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFacilityDatum.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.Datum = UIUtilsGen.GetComboBoxValueInt(cmbFacilityDatum)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbFacilityMethod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFacilityMethod.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.Method = UIUtilsGen.GetComboBoxValueInt(cmbFacilityMethod)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbFacilityLocationType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFacilityLocationType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.LocationType = UIUtilsGen.GetComboBoxValueInt(cmbFacilityLocationType)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dtFacilityPowerOff_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtFacilityPowerOff.ValueChanged
    '    If bolLoading Then Exit Sub
    '    Try
    '        UIUtilsGen.ToggleDateFormat(Me.dtFacilityPowerOff)
    '        pOwn.Facilities.DatePowerOff = UIUtilsGen.GetDatePickerValue(dtFacilityPowerOff)
    '    Catch ex As Exception
    '        ShowError(ex)
    '    End Try
    'End Sub

    Private Sub txtFacilityAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.TextChanged
        If bolLoading Then Exit Sub
        Try
            If txtFacilityAddress.Tag > 0 Then
                pOwn.Facilities.AddressID = Integer.Parse(Trim(txtFacilityAddress.Tag))
            Else
                pOwn.Facilities.AddressID = 0
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.DoubleClick
        Try
            AddressForm = New Address(UIUtilsGen.EntityTypes.Facility, pOwn.Facilities.FacilityAddresses, "Facility", pOwn.AddressId, UIUtilsGen.ModuleID.Registration)
            AddressForm.ShowFIPS = False
            AddressForm.ShowDialog()
            ' update txtfacilityaddress text
            pOwn.Facilities.FacilityAddresses.Retrieve(pOwn.Facilities.AddressID)
            txtFacilityAddress.Text = UIUtilsGen.FormatAddress(pOwn.Facilities.FacilityAddresses, True)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityAddress_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFacilityAddress.KeyUp
        Try
            If e.KeyCode = Keys.Enter Then
                txtFacilityAddress_DoubleClick(sender, New System.EventArgs)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtFacilityAddress_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.Enter
        Try
            If txtFacilityAddress.Text = String.Empty Then
                txtFacilityAddress_DoubleClick(sender, e)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub lblFacilityLicenseeSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFacilityLicenseeSearch.Click
        Try
            strFromCompanySearch = "FACILITY"
            oCompanySearch = New CompanySearch
            oCompanySearch.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        Finally
            oCompanySearch = Nothing
        End Try
    End Sub

    Private Sub lnkLblNextFac_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextFac.LinkClicked
        Try
            If Not bolAddFacility Then
                PopulateFacility(GetPrevNextFacility(nFacilityID, True))
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub lnkLblPrevFacility_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevFacility.LinkClicked
        Try
            If Not bolAddFacility Then
                PopulateFacility(GetPrevNextFacility(nFacilityID, False))
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnExpand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpand.Click
        Try
            If btnExpand.Text = "Expand All" Then
                ExpandAll(True, dgPipesAndTanks, btnExpand)
            Else
                ExpandAll(False, dgPipesAndTanks, btnExpand)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnExpandTP2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandTP2.Click
        Try
            If btnExpandTP2.Text = "Expand All" Then
                ExpandAll(True, dgPipesAndTanks2, btnExpandTP2)
            Else
                ExpandAll(False, dgPipesAndTanks2, btnExpandTP2)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnFacilitySave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacilitySave.Click
        Dim nUpcomingInstallRegActivity As Integer = -1
        Dim nSigReqActivity As Integer = -1
        Try
            If pOwn.Facilities.UpcomingInstallation And Not pOwn.Facilities.UpcomingInstallationOriginal Then
                nUpcomingInstallRegActivity = 1
                pOwn.Facilities.UpcomingInstallation = False
            Else
                If Date.Compare(pOwn.Facilities.UpcomingInstallationDate, dtNullDate) = 0 And Date.Compare(pOwn.Facilities.UpcomingInstallationDateOriginal, dtNullDate) <> 0 Then
                    nUpcomingInstallRegActivity = 0
                End If
            End If
            ' need to place facility in register mode only on signature changed to false
            ' and has tanks and facility update
            If Not pOwn.Facilities.SignatureOnNF And pOwn.Facilities.SignatureOnNFOriginal Then
                If Not dgPipesAndTanks.Rows Is Nothing Then
                    If dgPipesAndTanks.Rows.Count > 0 Then
                        nSigReqActivity = 1
                    End If
                End If
            ElseIf bolAddFacility And Not pOwn.Facilities.SignatureOnNF Then
                nSigReqActivity = 1
            ElseIf pOwn.Facilities.SignatureOnNF And Not pOwn.Facilities.SignatureOnNFOriginal Then
                nSigReqActivity = 0
            End If

            Dim success As Boolean = False
            If pOwn.Facilities.ID <= 0 Then
                pOwn.Facilities.CreatedBy = MC.AppUser.ID
            Else
                pOwn.Facilities.ModifiedBy = MC.AppUser.ID
            End If
            returnVal = String.Empty

            success = pOwn.Facilities.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)

            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If success Then
                ' MsgBox("Facility Saved Successfully")
                ' if new facility, update owner's facility grid
                If nFacilityID <= 0 Then
                    bolAddFacility = False
                    If nSigReqActivity <> -1 Then
                        If pOwn.Facilities.SignatureOnNF = False Then nSigReqActivity = 1
                    End If
                    PopulateOwnerFacs(nOwnerID)
                End If
                nFacilityID = pOwn.Facilities.ID

                Dim LocalUserSettings As Microsoft.Win32.Registry
                Dim conSQLConnection As New SqlConnection
                Dim cmdSQLCommand As New SqlCommand
                Dim aReader As SqlDataReader

                conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")

                conSQLConnection.Open()
                cmdSQLCommand.Connection = conSQLConnection

                cmdSQLCommand.CommandText = "select * from tblReg_AssessDate where FacilityID = " + lblFacilityIDValue.Text
                aReader = cmdSQLCommand.ExecuteReader()

                If aReader.HasRows() Then
                    cmdSQLCommand.CommandText = "update tblReg_AssessDate set AssessDate = " + IIf(dtPickAssess.Checked, "'" + dtPickAssess.Value.ToShortDateString.ToString + "'", "NULL") + " where FacilityID = " + nFacilityID.ToString
                Else
                    cmdSQLCommand.CommandText = "insert into tblReg_AssessDate values(" + nFacilityID.ToString + "," + IIf(dtPickAssess.Checked, "'" + dtPickAssess.Value.ToShortDateString.ToString + "'", "NULL") + ")"
                End If
                aReader.Close()
                cmdSQLCommand.ExecuteNonQuery()
                conSQLConnection.Close()
                cmdSQLCommand.Dispose()
                conSQLConnection.Dispose()

                MsgBox("Facility Saved Successfully")
                If nUpcomingInstallRegActivity = 1 Then
                    PutRegistrationActivity(nFacilityID.ToString, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.UpComingInstall)
                ElseIf nUpcomingInstallRegActivity = 0 Then
                    DeleteRegistrationActivity(nFacilityID.ToString, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.UpComingInstall)
                End If
                If nSigReqActivity = 1 Then
                    PutRegistrationActivity(nFacilityID.ToString, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.SignatureRequired)
                ElseIf nSigReqActivity = 0 Then
                    DeleteRegistrationActivity(nFacilityID.ToString, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.SignatureRequired)
                End If

                PopulateFacility(nFacilityID)
                UIUtilsGen.UpdateOwnerCAPInfo(Me, pOwn)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityCancel.Click
        Try
            If Not pOwn.Facilities Is Nothing Then
                pOwn.Facilities.Reset()
                PopulateFacility(nFacilityID)
                'If pOwn.Facilities.ID > 0 Then
                '    Me.PopulateFacilityInfo(pOwn.Facilities.ID)
                'End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnDeleteFacility_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeleteFacility.Click
        Try
            If MessageBox.Show("Are you sure you wish to DELETE the Facility?", "DELETE FACILITY", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                Exit Sub
            End If
            If pOwn.Facilities.IsDirty Then
                If MessageBox.Show("There are unsaved changes.Do you want to save the changes before delete? ", "Caption", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    pOwn.Facilities.Reset()
                End If
            End If
            If Not dgPipesAndTanks.Rows Is Nothing Then
                If dgPipesAndTanks.Rows.Count > 0 Then
                    MsgBox("The Specified facility has associated Tank(s). Delete Tank(s) before deleting the facility")
                    Exit Sub
                End If
            End If
            pOwn.Facilities.Deleted = True
            ' need to save only if existing facility
            If pOwn.Facilities.ID > 0 Then
                pOwn.Facilities.ModifiedBy = MC.AppUser.ID
                returnVal = String.Empty
                If Not pOwn.Facilities.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID, True) Then
                    Exit Sub
                End If

                ' delete registration activity
                CheckRegHeader(pOwn.ID)
                If oReg.ID > 0 Then
                    oReg.Activity.Col = oReg.Activity.RetrieveByRegID(oReg.ID)
                    For Each oregActivity As MUSTER.Info.RegistrationActivityInfo In oReg.Activity.Col.Values
                        If oregActivity.EntityId = pOwn.Facilities.ID And oregActivity.EntityType = UIUtilsGen.EntityTypes.Facility Then
                            oregActivity.Processed = True
                        End If
                    Next
                    oReg.Save()
                End If

                Dim nextFac As Integer = GetPrevNextFacility(nFacilityID, True)
                If nextFac <> nFacilityID Then
                    ' refresh facility grid on owner tab
                    UIUtilsGen.PopulateOwnerFacilities(pOwn, Me, nOwnerID)
                    PopulateFacility(nextFac)

                    ' update owner's facility grid
                    PopulateOwnerFacs(nOwnerID)
                Else
                    pOwn.Facilities = New MUSTER.BusinessLogic.pFacility
                    tbCntrlRegistration.SelectedTab = tbPageOwnerDetail
                End If
            Else
                pOwn.Facilities.Remove(pOwn.Facilities.ID)
                pOwn.Facilities = New MUSTER.BusinessLogic.pFacility
                tbCntrlRegistration.SelectedTab = tbPageOwnerDetail
                UIUtilsGen.UpdateOwnerCAPInfo(Me, pOwn)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnAddTank_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddTank.Click
        Try
            If nFacilityID <= 0 Then
                MsgBox("Please complete adding facility first.")
                Exit Sub
            End If
            bolAddTank = True
            'tbCntrlRegistration.Tag = tbCntrlTank.Name
            ResetTankID()
            ShowHideTankPipeScreen(True, False)
            If tbCntrlRegistration.SelectedTab.Name = tbPageManageTank.Name Then
                SetupTabs()
            Else
                tbCntrlRegistration.SelectedTab = tbPageManageTank
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnAddTank2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddTank2.Click
        btnAddTank_Click(sender, e)
    End Sub

    Private Sub LinkLblCAPSignupFac_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLblCAPSignupFac.LinkClicked
        Dim frmCAP As CAPSignUp
        Try
            frmCAP = New CAPSignUp(pOwn)
            frmCAP.nOwnerID = nOwnerID
            frmCAP.nFacilityID = nFacilityID
            frmCAP.txtOwner.Text = Me.txtOwnerName.Text + " (" + Me.lblOwnerIDValue.Text + ")"
            'frmCAP.txtOwner.Tag = Me.txtOwnerName.Text
            Me.Tag = "0"
            frmCAP.CallingForm = Me
            frmCAP.ShowDialog()
            If Me.Tag = "1" Then
                ' clearing the collection
                pOwn.OwnerInfo.facilityCollection = New MUSTER.Info.FacilityCollection
                pOwn.Facilities = New MUSTER.BusinessLogic.pFacility
                pTank = New MUSTER.BusinessLogic.pTank
                pPipe = New MUSTER.BusinessLogic.pPipe
                PopulateFacility(nFacilityID)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region

#Region "Tank"
    Friend Sub PopulateTank(ByVal tnkID As Integer)
        Try
            ' Retreive info from tblReg_Prohibition
            Dim LocalUserSettings As Microsoft.Win32.Registry
            Dim conSQLConnection As New SqlConnection
            Dim cmdSQLCommand As New SqlCommand
            Dim prohibitionReader As SqlDataReader
            conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")
            conSQLConnection.Open()
            cmdSQLCommand.Connection = conSQLConnection
            cmdSQLCommand.CommandText = "select * from tblReg_Prohibition where Facility_id = " + nFacilityID.ToString + " and tank_id = " + tnkID.ToString
            prohibitionReader = cmdSQLCommand.ExecuteReader()
            autoChange = True
            If prohibitionReader.HasRows() Then
                Me.chkTankProhibition.Checked = True
            Else
                Me.chkTankProhibition.Checked = False
            End If
            autoChange = False
            prohibitionReader.Close()
            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()
            ShowHideTankPipeScreen(True, False)

            SetTankRow(tnkID)

            ResetCompartmentNumber()
            nTankID = tnkID

            pTank = pOwn.Facilities.FacilityTanks
            pTank.RetrieveTank(tnkID)

            If pTank.TankId <= 0 Then
                nTankID = pTank.TankId
                pTank.Compartments.Retrieve(pTank.TankInfo, tnkID, False)

                ClearTankForm()

                FillTankForm()

                Me.Text = "Registration - Manage Tank (New) - " & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"

                chkBoxReplacementTank.Checked = False
                chkBoxReplacementTank.Visible = True
                CommentsMaintenance(, , True, True)
            Else
                chkBoxReplacementTank.Checked = False
                chkBoxReplacementTank.Visible = False

                ClearTankForm()

                FillTankForm()

                ' goto pipe is enabled only if tank has pipes in the grid
                If Not ugTankRow.ChildBands(0) Is Nothing Then
                    If Not ugTankRow.ChildBands(0).Rows Is Nothing Then
                        If ugTankRow.ChildBands(0).Rows.Count > 0 Then
                            btnToPipe.Enabled = True
                        End If
                    End If
                End If

                Me.Text = "Registration - Manage Tank (" + pTank.TankIndex.ToString + ") - " & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
                MC.FlagsChanged(nFacilityID, UIUtilsGen.EntityTypes.Facility, "Registration", "NONE", nTankID, UIUtilsGen.EntityTypes.Tank)
                CommentsMaintenance(, , True)
            End If
            If Me.cmbTanksubstance.SelectedValue = 314 Then ' used oil
                UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionInstalled, dtNullDate)
                UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionLastTested, dtNullDate)
                UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionLastInspected, dtNullDate)
                UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionInstalled, dtNullDate)
                Me.dtPickSpillPreventionInstalled.Enabled = False
                Me.dtPickSpillPreventionLastTested.Enabled = False
                Me.dtPickOverfillPreventionLastInspected.Enabled = False
                Me.dtPickOverfillPreventionInstalled.Enabled = False
                '    Me.dtPickSpillPreventionInstalled.Value = "1900-01-01"
                '    Me.dtPickSpillPreventionLastTested.Value = "1900-01-01"
                '    Me.dtPickOverfillPreventionInstalled.Value = "1900-01-01"
                '    Me.dtPickOverfillPreventionLastInspected.Value = "1900-01-01"
            Else
                If Not pTank.SmallDelivery Then
                    Me.dtPickSpillPreventionInstalled.Enabled = True
                    Me.dtPickSpillPreventionLastTested.Enabled = True
                    Me.dtPickOverfillPreventionLastInspected.Enabled = True
                    Me.dtPickOverfillPreventionInstalled.Enabled = True
                Else
                    Me.dtPickSpillPreventionInstalled.Enabled = False
                    Me.dtPickSpillPreventionLastTested.Enabled = False
                    Me.dtPickOverfillPreventionLastInspected.Enabled = False
                    Me.dtPickOverfillPreventionInstalled.Enabled = False
                    '    Me.dtPickSpillPreventionInstalled.Value = "1900-01-01"
                    '    Me.dtPickSpillPreventionLastTested.Value = "1900-01-01"
                    '    Me.dtPickOverfillPreventionInstalled.Value = "1900-01-01"
                    '    Me.dtPickOverfillPreventionLastInspected.Value = "1900-01-01"
                End If
            End If
            If Me.cmbTankReleaseDetection.SelectedValue = 336 Then 'Automatic Tank Gauging - 336
                Me.dtPickATGLastInspected.Enabled = True
            Else
                UIUtilsGen.SetDatePickerValue(dtPickATGLastInspected, dtNullDate)
                Me.dtPickATGLastInspected.Enabled = False
            End If
            If Me.cmbTankReleaseDetection.SelectedValue = 339 Then 'Electronic Interstitial Monitoring -339
                Me.dtPickElectronicDeviceInspected.Enabled = True
            Else
                UIUtilsGen.SetDatePickerValue(dtPickElectronicDeviceInspected, dtNullDate)
                Me.dtPickElectronicDeviceInspected.Enabled = False
            End If
            '  If Thank Release Detection = like ‘%Interstitial Monitoring%’, disable dtPickSecondaryContainmentLastInspected
            '  If (Me.dtPickTankInstalled.Value < Convert.ToDateTime("10-01-2008")) Or (Me.cmbTankReleaseDetection.SelectedValue = 339) Or (Me.cmbTankReleaseDetection.SelectedValue = 343) Then
            ' Me.dtPickSecondaryContainmentLastInspected.Enabled = False
            '  End If
            If Me.cmbTankReleaseDetection.SelectedValue = 343 Then 'Visual Interstitial Monitoring
                Me.dtPickSecondaryContainmentLastInspected.Enabled = True
            Else
                UIUtilsGen.SetDatePickerValue(dtPickSecondaryContainmentLastInspected, dtNullDate)
                Me.dtPickSecondaryContainmentLastInspected.Enabled = False
            End If

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub SetTankRow(ByVal tnkID As Integer)
        Try
            If ugTankRow Is Nothing Then
                For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In dgPipesAndTanks.Rows
                    If ugRow.Cells("TANK ID").Value = tnkID Then
                        ugTankRow = ugRow
                        Exit For
                    End If
                Next
            Else
                If ugTankRow.Cells("TANK ID").Value <> tnkID Then
                    ugTankRow = Nothing
                    SetTankRow(tnkID)
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub SetTankSaveCancel(ByVal bolValue As Boolean)
        btnTankSave.Enabled = bolValue
        btnTankCancel.Enabled = bolValue
        If nTankID > 0 Then
            btnDeleteTank.Enabled = True
        ElseIf bolValue Then
            btnDeleteTank.Enabled = True
        End If
    End Sub
    Private Sub ClearTankForm()
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True

            lblTankIDValue2.Text = String.Empty
            lblTankIDValue.Text = String.Empty

            lblOwnerLastEditedBy.Text = "Last Edited By : "
            lblOwnerLastEditedOn.Text = "Last Edited On : "

            cmbTankStatus.DataSource = Nothing
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickTankInstalled)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickDatePlacedInService)
            cmbTankManufacturer.DataSource = Nothing

            cmbTankStatus.Enabled = False
            dtPickTankInstalled.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickTankInstalled)
            dtPickDatePlacedInService.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickDatePlacedInService)
            cmbTankManufacturer.Enabled = False

            ClearCompartmentArea()

            cmbTankMaterial.DataSource = Nothing
            cmbTankOptions.DataSource = Nothing
            cmbTankCPType.DataSource = Nothing
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCPInstalled)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCPLastTested)
            chkEmergencyPower.Checked = False
            chkDeliveriesLimited.Checked = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickInteriorLiningInstalled)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickLastInteriorLinningInspection)
            cmbTankOverfillProtectionType.DataSource = Nothing
            chkBxSpillProtected.Checked = False
            chkBxTightfillAdapters.Checked = False
            chkOverFilledProtected.Checked = False

            cmbTankMaterial.Enabled = False
            cmbTankOptions.Enabled = False
            cmbTankCPType.Enabled = False
            dtPickCPInstalled.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCPInstalled)
            dtPickCPLastTested.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCPLastTested)
            chkEmergencyPower.Enabled = False
            chkDeliveriesLimited.Enabled = False
            dtPickInteriorLiningInstalled.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickInteriorLiningInstalled)
            dtPickLastInteriorLinningInspection.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickLastInteriorLinningInspection)
            cmbTankOverfillProtectionType.Enabled = False
            chkBxSpillProtected.Enabled = False
            chkBxTightfillAdapters.Enabled = False
            chkOverFilledProtected.Enabled = False

            cmbTankReleaseDetection.DataSource = Nothing
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickTankTightnessTest)
            chkTankDrpTubeInvControl.Checked = False

            cmbTankReleaseDetection.Enabled = False
            dtPickTankTightnessTest.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickTankTightnessTest)
            chkTankDrpTubeInvControl.Enabled = False

            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickTankInstallerSigned)
            txtLicensee.Text = String.Empty
            txtTankCompany.Text = String.Empty

            dtPickTankInstallerSigned.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickTankInstallerSigned)
            txtLicensee.Enabled = False
            txtTankCompany.Enabled = False
            lblLicenseeSearch.Enabled = False

            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickLastUsed)
            lblDateTankClosureRecvValue.Text = String.Empty
            lblClosuredate.Text = String.Empty
            cmbTankClosureType.DataSource = Nothing
            cmbTankInertFill.DataSource = Nothing

            dtPickLastUsed.Enabled = False
            lblDateTankClosureRecvValue.Enabled = False
            lblClosuredate.Enabled = False
            cmbTankClosureType.Enabled = False
            cmbTankInertFill.Enabled = False

            btnTankSave.Enabled = False
            btnTankCancel.Enabled = False
            btnCopyTankProfileToNew.Enabled = False
            btnDeleteTank.Enabled = False
            btnToPipe.Enabled = False
            btnTankComments.Enabled = False
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub ClearCompartmentArea()
        Try
            chkTankCompartment.Checked = False

            txtTankCapacity.Text = String.Empty
            txtTankCompartmentNumber.Text = String.Empty
            dGridCompartments.DataSource = Nothing

            txtTankCapacity.Enabled = False
            txtTankCompartmentNumber.Enabled = False
            dGridCompartments.Enabled = False

            lblTankCapacity.Visible = False
            txtTankCapacity.Visible = False
            lblTankCompartmentNumber.Visible = False
            txtTankCompartmentNumber.Visible = False
            dGridCompartments.Visible = False

            pnlNonCompProperties.Visible = True
            txtNonCompTankCapacity.Text = String.Empty
            cmbTanksubstance.DataSource = Nothing
            cmbTankFuelType.DataSource = Nothing
            cmbTankCercla.DataSource = Nothing
            cmbTankCerclaDesc.DataSource = Nothing

            txtNonCompTankCapacity.Enabled = True
            cmbTanksubstance.Enabled = True
            cmbTankFuelType.Enabled = True
            cmbTankCercla.Enabled = False
            cmbTankCerclaDesc.Enabled = False

            ttCERCLA.AutoPopDelay = 15000 ' set lenth tip is visible to 15 seconds
            ttCERCLA.SetToolTip(lblCERCLAtt, "None")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FillTankForm()
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            If nTankID <= 0 Then
                lblTankIDValue2.Text = String.Empty
                lblTankIDValue.Text = String.Empty

                lblOwnerLastEditedBy.Text = "Last Edited By : "
                lblOwnerLastEditedOn.Text = "Last Edited On : "

                ' Status
                EnableTankStatus(True)
                PopulateTankStatus("ADD", pTank.DateLastUsed)
                UIUtilsGen.SetComboboxItemByValue(cmbTankStatus, pTank.TankStatus)
                cmbTankStatus.SelectedIndex = -1
                If cmbTankStatus.SelectedIndex <> -1 Then
                    cmbTankStatus.SelectedIndex = -1
                End If

                FillTankFormFields(True, True)
            Else
                bolLoading = True

                lblTankIDValue2.Text = pTank.TankIndex.ToString
                lblTankIDValue.Text = pTank.TankIndex.ToString

                lblOwnerLastEditedBy.Text = "Last Edited By : " & IIf(pTank.ModifiedBy = String.Empty, pTank.CreatedBy, pTank.ModifiedBy)
                If Date.Compare(pTank.ModifiedOn, dtNullDate) = 0 Then
                    lblOwnerLastEditedOn.Text = "Last Edited On : " & pTank.CreatedOn.ToString
                Else
                    lblOwnerLastEditedOn.Text = "Last Edited On : " & pTank.ModifiedOn.ToString
                End If

                ' Status
                EnableTankStatus(True)
                PopulateTankStatus("EDIT", pTank.DateLastUsed)
                UIUtilsGen.SetComboboxItemByValue(cmbTankStatus, pTank.TankStatus)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankStatus, pTank.TankStatus, "PROPERTY_ID", "=") Then
                    pTank.TankStatus = 0
                End If
                If pTank.TankStatus <= 0 Then
                    cmbTankStatus.SelectedIndex = -1
                    If cmbTankStatus.SelectedIndex <> -1 Then
                        cmbTankStatus.SelectedIndex = -1
                    End If
                End If

                FillTankFormFields(True, True)
            End If
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub FillTankFormFields(ByVal populateCombo As Boolean, ByVal populateComps As Boolean)
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            Dim nTankCapacity As Integer = 0
            Dim bolDisableOverfillType As Boolean = False
            If cmbTankStatus.Text <> String.Empty Or pTank.TankStatus > 0 Then
                dtPickTankInstalled.Enabled = True
                UIUtilsGen.SetDatePickerValue(dtPickTankInstalled, pTank.DateInstalledTank)

                dtPickDatePlacedInService.Enabled = True
                UIUtilsGen.SetDatePickerValue(dtPickDatePlacedInService, pTank.PlacedInServiceDate)

                cmbTankManufacturer.Enabled = True
                PopulateTankManufacturer()
                UIUtilsGen.SetComboboxItemByValue(cmbTankManufacturer, pTank.TankManufacturer)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankManufacturer, pTank.TankManufacturer, "PROPERTY_ID", "=") Then
                    pTank.TankManufacturer = 0
                End If
                If pTank.TankManufacturer <= 0 Then
                    cmbTankManufacturer.SelectedIndex = -1
                    If cmbTankManufacturer.SelectedIndex <> -1 Then
                        cmbTankManufacturer.SelectedIndex = -1
                    End If
                End If

                If populateComps Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, nTankID, False)
                    If pTank.Compartment Then
                        PopulateCompartments(nTankID)
                        If txtTankCapacity.Text = String.Empty Then
                            nTankCapacity = 0
                        Else
                            nTankCapacity = Integer.Parse(txtTankCapacity.Text)
                        End If
                    Else
                        PopulateCompartment(pTank.Compartments.COMPARTMENTNumber)
                        If txtNonCompTankCapacity.Text = String.Empty Then
                            nTankCapacity = 0
                        Else
                            nTankCapacity = Integer.Parse(txtNonCompTankCapacity.Text)
                        End If
                    End If
                Else
                    If pTank.Compartment Then
                        If txtTankCapacity.Text = String.Empty Then
                            nTankCapacity = 0
                        Else
                            nTankCapacity = Integer.Parse(txtTankCapacity.Text)
                        End If
                    Else
                        If txtNonCompTankCapacity.Text = String.Empty Then
                            nTankCapacity = 0
                        Else
                            nTankCapacity = Integer.Parse(txtNonCompTankCapacity.Text)
                        End If
                    End If
                End If
                ' #2889
                If Not pTank.Compartment Then
                    If pTank.Compartments.Substance = 314 Then
                        bolDisableOverfillType = True
                    End If
                End If

                ' TankMatDesc = Tank Material
                ' TankModDesc = Secondary Tank Option
                cmbTankMaterial.Enabled = True
                PopulateTankMaterial()
                UIUtilsGen.SetComboboxItemByValue(cmbTankMaterial, pTank.TankMatDesc)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankMaterial, pTank.TankMatDesc, "PROPERTY_ID", "=") Then
                    pTank.TankMatDesc = 0
                End If
                If pTank.TankMatDesc <= 0 Then
                    cmbTankMaterial.SelectedIndex = -1
                    If cmbTankMaterial.SelectedIndex <> -1 Then
                        cmbTankMaterial.SelectedIndex = -1
                    End If
                End If

                cmbTankOptions.Enabled = True
                PopulateTankOptions(pTank.TankMatDesc)
                UIUtilsGen.SetComboboxItemByValue(cmbTankOptions, pTank.TankModDesc)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankOptions, pTank.TankModDesc, "PROPERTY_ID", "=") Then
                    pTank.TankModDesc = 0
                End If
                If pTank.TankModDesc <= 0 Then
                    cmbTankOptions.SelectedIndex = -1
                    If cmbTankOptions.SelectedIndex <> -1 Then
                        cmbTankOptions.SelectedIndex = -1
                    End If
                End If

                ' Enable Field              Condition
                ' Tank CP Type              Tank Mod Desc (Tank Sec Option) like 'Cathodically Protected'
                ' Tank CP Installed         Tank Mod Desc (Tank Sec Option) like 'Cathodically Protected'
                ' Tank CP Last Tested       Tank Mod Desc (Tank Sec Option) like 'Cathodically Protected'
                ' Lined Interior Install    Tank Mod Desc (Tank Sec Option) like 'Lined'
                ' Lined Interior Inspect    Tank Mod Desc (Tank Sec Option) like 'Lined'
                If cmbTankOptions.Text.IndexOf("Cathodically Protected") > -1 Then
                    cmbTankCPType.Enabled = True
                    PopulateTankCPType()
                    UIUtilsGen.SetComboboxItemByValue(cmbTankCPType, pTank.TankCPType)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankCPType, pTank.TankCPType, "PROPERTY_ID", "=") Then
                        pTank.TankCPType = 0
                    End If
                    If pTank.TankCPType <= 0 Then
                        cmbTankCPType.SelectedIndex = -1
                        If cmbTankCPType.SelectedIndex <> -1 Then
                            cmbTankCPType.SelectedIndex = -1
                        End If
                    End If

                    dtPickCPInstalled.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickCPInstalled, pTank.TCPInstallDate)

                    dtPickCPLastTested.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickCPLastTested, pTank.LastTCPDate)

                    dtPickInteriorLiningInstalled.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickInteriorLiningInstalled)
                    pTank.LinedInteriorInstallDate = dtNullDate

                    dtPickLastInteriorLinningInspection.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickLastInteriorLinningInspection)
                    pTank.LinedInteriorInspectDate = dtNullDate
                ElseIf cmbTankOptions.Text.IndexOf("Lined") > -1 Then
                    dtPickInteriorLiningInstalled.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickInteriorLiningInstalled, pTank.LinedInteriorInstallDate)

                    dtPickLastInteriorLinningInspection.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickLastInteriorLinningInspection, pTank.LinedInteriorInspectDate)

                    cmbTankCPType.Enabled = False
                    cmbTankCPType.DataSource = Nothing
                    pTank.TankCPType = 0

                    dtPickCPInstalled.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCPInstalled)
                    pTank.TCPInstallDate = dtNullDate

                    dtPickCPLastTested.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCPLastTested)
                    pTank.LastTCPDate = dtNullDate
                Else
                    cmbTankCPType.Enabled = False
                    cmbTankCPType.DataSource = Nothing
                    pTank.TankCPType = 0

                    dtPickCPInstalled.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCPInstalled)
                    pTank.TCPInstallDate = dtNullDate

                    dtPickCPLastTested.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCPLastTested)
                    pTank.LastTCPDate = dtNullDate

                    dtPickInteriorLiningInstalled.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickInteriorLiningInstalled)
                    pTank.LinedInteriorInstallDate = dtNullDate

                    dtPickLastInteriorLinningInspection.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickLastInteriorLinningInspection)
                    pTank.LinedInteriorInspectDate = dtNullDate
                End If

                chkEmergencyPower.Enabled = True
                chkEmergencyPower.Checked = pTank.TankEmergen

                chkDeliveriesLimited.Enabled = True
                chkDeliveriesLimited.Checked = pTank.SmallDelivery
                If (chkDeliveriesLimited.Checked) Then
                    Me.dtPickOverfillPreventionInstalled.Enabled = False
                    Me.dtPickOverfillPreventionLastInspected.Enabled = False
                    Me.dtPickSpillPreventionInstalled.Enabled = False
                    Me.dtPickSpillPreventionLastTested.Enabled = False
                    '    Me.dtPickSpillPreventionInstalled.Value = "1900-01-01"
                    '    Me.dtPickSpillPreventionLastTested.Value = "1900-01-01"
                    '    Me.dtPickOverfillPreventionInstalled.Value = "1900-01-01"
                    '    Me.dtPickOverfillPreventionLastInspected.Value = "1900-01-01"
                Else
                    Me.dtPickOverfillPreventionInstalled.Enabled = True
                    Me.dtPickOverfillPreventionLastInspected.Enabled = True
                    Me.dtPickSpillPreventionInstalled.Enabled = True
                    Me.dtPickSpillPreventionLastTested.Enabled = True

                End If


                cmbTankOverfillProtectionType.Enabled = True
                PopulateTankOverfillProtectionType()
                UIUtilsGen.SetComboboxItemByValue(cmbTankOverfillProtectionType, pTank.OverFillType)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankOverfillProtectionType, pTank.OverFillType, "PROPERTY_ID", "=") Then
                    pTank.OverFillType = 0
                End If
                If pTank.OverFillType <= 0 Then
                    cmbTankOverfillProtectionType.SelectedIndex = -1
                    If cmbTankOverfillProtectionType.SelectedIndex <> -1 Then
                        cmbTankOverfillProtectionType.SelectedIndex = -1
                    End If
                End If

                chkBxSpillProtected.Enabled = True
                chkBxSpillProtected.Checked = pTank.SpillInstalled

                chkBxTightfillAdapters.Enabled = True
                chkBxTightfillAdapters.Checked = pTank.TightFillAdapters

                chkOverFilledProtected.Enabled = True
                chkOverFilledProtected.Checked = pTank.OverFillInstalled

                ' Need to enable Release Detection (Tank LD) Only if there is some value in Tank Mod Desc
                ' cause it returns no values if = 0
                If pTank.TankModDesc > 0 Then
                    cmbTankReleaseDetection.Enabled = True
                    PopulateTankReleaseDetection(pTank.TankModDesc, nTankCapacity)
                    UIUtilsGen.SetComboboxItemByValue(cmbTankReleaseDetection, pTank.TankLD)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankReleaseDetection, pTank.TankLD, "PROPERTY_ID", "=") Then
                        pTank.TankLD = 0
                    End If
                    If pTank.TankLD <= 0 Then
                        cmbTankReleaseDetection.SelectedIndex = -1
                        If cmbTankReleaseDetection.SelectedIndex <> -1 Then
                            cmbTankReleaseDetection.SelectedIndex = -1
                        End If
                    End If
                End If

                If (chkEmergencyPower.Checked = False And cmbTankReleaseDetection.Text.Trim.IndexOf("Deferred") > -1) Or _
                    (nTankCapacity > 2000 And cmbTankReleaseDetection.Text.Trim.IndexOf("Manual Tank Gauging") > -1) Then
                    pTank.TankLD = 0
                    cmbTankReleaseDetection.SelectedIndex = -1
                    If cmbTankReleaseDetection.SelectedIndex <> -1 Then
                        cmbTankReleaseDetection.SelectedIndex = -1
                    End If
                End If
                If chkEmergencyPower.Checked Then
                    cmbTankReleaseDetection.SelectedIndex = 0
                    pTank.TankLD = UIUtilsGen.GetComboBoxValueInt(cmbTankReleaseDetection)
                End If

                If chkDeliveriesLimited.Checked Then
                    '---- uncheck / clear spill, overfill, tight fill adapter and overfill type
                    Me.dtPickOverfillPreventionInstalled.Enabled = False
                    Me.dtPickOverfillPreventionLastInspected.Enabled = False
                    Me.dtPickSpillPreventionInstalled.Enabled = False
                    Me.dtPickSpillPreventionLastTested.Enabled = False
                    '  Me.dtPickSpillPreventionInstalled.Value = "1900-01-01"
                    '  Me.dtPickSpillPreventionLastTested.Value = "1900-01-01"
                    '  Me.dtPickOverfillPreventionInstalled.Value = "1900-01-01"
                    '  Me.dtPickOverfillPreventionLastInspected.Value = "1900-01-01"

                    chkBxSpillProtected.Checked = False
                    pTank.SpillInstalled = False

                    chkBxTightfillAdapters.Checked = False
                    pTank.TightFillAdapters = False

                    chkOverFilledProtected.Checked = False
                    pTank.OverFillInstalled = False

                    cmbTankOverfillProtectionType.SelectedIndex = -1
                    If cmbTankOverfillProtectionType.SelectedIndex <> -1 Then
                        cmbTankOverfillProtectionType.SelectedIndex = -1
                    End If
                    pTank.OverFillType = 0

                    '----- disable spill, overfill, tight fill adapter and overfill type
                    chkBxSpillProtected.Enabled = False
                    chkOverFilledProtected.Enabled = False
                    cmbTankOverfillProtectionType.Enabled = False
                    chkBxTightfillAdapters.Enabled = False
                Else
                    chkBxSpillProtected.Enabled = True
                    chkOverFilledProtected.Enabled = True
                    cmbTankOverfillProtectionType.Enabled = True
                    chkBxTightfillAdapters.Enabled = True

                    chkBxSpillProtected.Checked = pTank.SpillInstalled
                    chkOverFilledProtected.Checked = pTank.OverFillInstalled
                    chkBxTightfillAdapters.Checked = pTank.TightFillAdapters
                    UIUtilsGen.SetComboboxItemByValue(cmbTankOverfillProtectionType, pTank.OverFillType)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankOverfillProtectionType, pTank.OverFillType, "PROPERTY_ID", "=") Then
                        pTank.OverFillType = 0
                    End If
                    If pTank.OverFillType <= 0 Then
                        cmbTankOverfillProtectionType.SelectedIndex = -1
                        If cmbTankOverfillProtectionType.SelectedIndex <> -1 Then
                            cmbTankOverfillProtectionType.SelectedIndex = -1
                        End If
                    End If
                End If

                If bolDisableOverfillType Then
                    cmbTankOverfillProtectionType.SelectedIndex = -1
                    If cmbTankOverfillProtectionType.SelectedIndex <> -1 Then
                        cmbTankOverfillProtectionType.SelectedIndex = -1
                    End If
                    pTank.OverFillType = 0
                    cmbTankOverfillProtectionType.Enabled = False
                End If

                ' Enable Field              Condition
                ' TTT Date                  Tank LD = Inventory Control / Precision Tightness Test
                ' Droptube for IC           Tank LD = Inventory Control / Precision Tightness Test (Automatically check true)
                If cmbTankReleaseDetection.Text.IndexOf("Inventory Control/Precision Tightness Testing") > -1 Then
                    dtPickTankTightnessTest.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickTankTightnessTest, pTank.TTTDate)
                    chkTankDrpTubeInvControl.Enabled = True
                    chkTankDrpTubeInvControl.Checked = True
                    pTank.DropTube = True
                Else
                    dtPickTankTightnessTest.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickTankTightnessTest)
                    pTank.TTTDate = dtNullDate
                    chkTankDrpTubeInvControl.Enabled = False
                    chkTankDrpTubeInvControl.Checked = False
                    pTank.DropTube = False
                End If

                FillTankInstallerOathLicensee()

                If Date.Compare(pTank.DateClosureReceived, dtNullDate) = 0 Then
                    lblDateTankClosureRecvValue.Text = String.Empty
                Else
                    lblDateTankClosureRecvValue.Text = pTank.DateClosureReceived.ToShortDateString
                End If
                If Date.Compare(pTank.DateClosed, dtNullDate) = 0 Then
                    lblClosuredate.Text = String.Empty
                Else
                    lblClosuredate.Text = pTank.DateClosed.ToShortDateString
                End If

                ' to populate the values. enabling and disabling is done after populating
                cmbTankClosureType.Enabled = True
                cmbTankInertFill.Enabled = True
                PopulateTankClosureType()
                PopulateTankInertFill()
                UIUtilsGen.SetDatePickerValue(dtPickLastUsed, pTank.DateLastUsed)

                UIUtilsGen.SetComboboxItemByValue(cmbTankClosureType, pTank.ClosureType)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankClosureType, pTank.ClosureType, "PROPERTY_ID", "=") Then
                    pTank.ClosureType = 0
                End If
                If pTank.ClosureType <= 0 Then
                    cmbTankClosureType.SelectedIndex = -1
                    If cmbTankClosureType.SelectedIndex <> -1 Then
                        cmbTankClosureType.SelectedIndex = -1
                    End If
                End If
                UIUtilsGen.SetComboboxItemByValue(cmbTankInertFill, pTank.InertMaterial)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankInertFill, pTank.InertMaterial, "PROPERTY_ID", "=") Then
                    pTank.InertMaterial = 0
                End If
                If pTank.InertMaterial <= 0 Then
                    cmbTankInertFill.SelectedIndex = -1
                    If cmbTankInertFill.SelectedIndex <> -1 Then
                        cmbTankInertFill.SelectedIndex = -1
                    End If
                End If

                If cmbTankStatus.Text.IndexOf("Currently In Use") > -1 Then
                    dtPickLastUsed.Enabled = False
                    cmbTankClosureType.Enabled = False
                    cmbTankInertFill.Enabled = False
                Else
                    dtPickLastUsed.Enabled = True
                    cmbTankClosureType.Enabled = True
                    cmbTankInertFill.Enabled = True
                End If
            End If

            If nTankID > 0 Then
                btnAddTank2.Enabled = True
                btnAddPipe.Enabled = True
                btnAddExistingPipe.Enabled = True
                btnCopyTankProfileToNew.Enabled = True
                btnTankComments.Enabled = True
                ' #
                Dim strAttachedPipeIDs As String = pTank.GetAttachedPipeIDs(nTankID)
                If strAttachedPipeIDs.Length > 0 Then
                    btnDetachPipes.Enabled = True
                Else
                    btnDetachPipes.Enabled = False
                End If
            Else
                btnAddTank2.Enabled = False
                btnAddPipe.Enabled = False
                btnAddExistingPipe.Enabled = False
                btnCopyTankProfileToNew.Enabled = False
                btnTankComments.Enabled = False
                btnToPipe.Enabled = False
                btnDetachPipes.Enabled = False
            End If

            ' Retreive info from tblReg_Prohibition
            Dim LocalUserSettings As Microsoft.Win32.Registry
            Dim conSQLConnection As New SqlConnection
            Dim cmdSQLCommand As New SqlCommand
            Dim prohibitionReader As SqlDataReader
            Dim tankPlusReader As SqlDataReader

            conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")
            conSQLConnection.Open()
            cmdSQLCommand.Connection = conSQLConnection
            cmdSQLCommand.CommandText = "select * from tblReg_Prohibition where Facility_id = " + nFacilityID.ToString
            prohibitionReader = cmdSQLCommand.ExecuteReader()

            If prohibitionReader.HasRows() Then
                Me.chkProhibition.Checked = True
            Else
                Me.chkProhibition.Checked = False
            End If

            prohibitionReader.Close()

            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()

            'Retreive tankPlus date info  -  added by Hua Cao on 09/09/2008
            If (pTank.TankId > 0) And (populateCombo = True) Then
                Dim ds As DataTable = pTank.GetDataTable("tblReg_TankPlus where TankID = " + pTank.TankId.ToString)
                If Not (ds Is Nothing) Then
                    If Not ds.Rows(0).Item("DateSpillPreventionInstalled") Is System.DBNull.Value Then
                        Me.dtPickSpillPreventionInstalled.Format = DateTimePickerFormat.Short
                        dtPickSpillPreventionInstalled.Value = ds.Rows(0).Item("DateSpillPreventionInstalled")
                        dtPickSpillPreventionInstalled.Checked = True
                    Else
                        UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionInstalled, dtNullDate)
                    End If
                    If Not ds.Rows(0).Item("DateOverfillPreventionInstalled") Is System.DBNull.Value Then
                        dtPickOverfillPreventionInstalled.Format = DateTimePickerFormat.Short
                        dtPickOverfillPreventionInstalled.Value = ds.Rows(0).Item("DateOverfillPreventionInstalled")
                        dtPickOverfillPreventionInstalled.Checked = True
                    Else
                        UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionInstalled, dtNullDate)
                    End If
                    If Not ds.Rows(0).Item("DateSpillPreventionLastTested") Is System.DBNull.Value Then
                        dtPickSpillPreventionLastTested.Format = DateTimePickerFormat.Short
                        dtPickSpillPreventionLastTested.Value = ds.Rows(0).Item("DateSpillPreventionLastTested")
                        dtPickSpillPreventionLastTested.Checked = True
                    Else
                        UIUtilsGen.SetDatePickerValue(Me.dtPickSpillPreventionLastTested, dtNullDate)
                    End If
                    If Not ds.Rows(0).Item("DateOverfillPreventionLastInspected") Is System.DBNull.Value Then
                        dtPickOverfillPreventionLastInspected.Format = DateTimePickerFormat.Short
                        dtPickOverfillPreventionLastInspected.Value = ds.Rows(0).Item("DateOverfillPreventionLastInspected")
                        dtPickOverfillPreventionLastInspected.Checked = True
                    Else
                        UIUtilsGen.SetDatePickerValue(Me.dtPickOverfillPreventionLastInspected, dtNullDate)
                    End If
                    If Not ds.Rows(0).Item("DateSecondaryContainmentLastInspected") Is System.DBNull.Value Then
                        dtPickSecondaryContainmentLastInspected.Format = DateTimePickerFormat.Short
                        dtPickSecondaryContainmentLastInspected.Value = ds.Rows(0).Item("DateSecondaryContainmentLastInspected")
                        dtPickSecondaryContainmentLastInspected.Checked = True
                    Else
                        UIUtilsGen.SetDatePickerValue(Me.dtPickSecondaryContainmentLastInspected, dtNullDate)
                    End If
                    If Not ds.Rows(0).Item("DateElectronicDeviceInspected") Is System.DBNull.Value Then
                        dtPickElectronicDeviceInspected.Format = DateTimePickerFormat.Short
                        dtPickElectronicDeviceInspected.Value = ds.Rows(0).Item("DateElectronicDeviceInspected")
                        dtPickElectronicDeviceInspected.Checked = True
                    Else
                        UIUtilsGen.SetDatePickerValue(Me.dtPickElectronicDeviceInspected, dtNullDate)
                    End If
                    If Not ds.Rows(0).Item("DateATGLastInspected") Is System.DBNull.Value Then
                        dtPickATGLastInspected.Format = DateTimePickerFormat.Short
                        dtPickATGLastInspected.Value = ds.Rows(0).Item("DateATGLastInspected")
                        dtPickATGLastInspected.Checked = True
                    Else
                        UIUtilsGen.SetDatePickerValue(Me.dtPickATGLastInspected, dtNullDate)
                    End If
                Else
                    UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionInstalled, dtNullDate)
                    UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionInstalled, dtNullDate)
                    UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionLastTested, dtNullDate)
                    UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionLastInspected, dtNullDate)
                    UIUtilsGen.SetDatePickerValue(dtPickSecondaryContainmentLastInspected, dtNullDate)
                    UIUtilsGen.SetDatePickerValue(dtPickElectronicDeviceInspected, dtNullDate)
                    UIUtilsGen.SetDatePickerValue(dtPickATGLastInspected, dtNullDate)
                End If
            End If
            SetTankSaveCancel(pTank.IsDirty)

        Catch ex As Exception
            Throw ex
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub FillTankInstallerOathLicensee()
        Try
            dtPickTankInstallerSigned.Enabled = True
            UIUtilsGen.SetDatePickerValue(dtPickTankInstallerSigned, pTank.DateSigned)

            txtLicensee.Enabled = True
            txtTankCompany.Enabled = True
            lblLicenseeSearch.Enabled = True

            pLicensee.Retrieve(pTank.LicenseedID)
            txtLicensee.Text = pLicensee.Licensee_name

            pCompany.Retrieve(pTank.ContractorID)
            txtTankCompany.Text = pCompany.COMPANY_NAME
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub EnableTankStatus(ByVal bolValue As Boolean)
        cmbTankStatus.Enabled = bolValue
    End Sub
    Private Sub PopulateTankStatus(ByVal Mode As String, Optional ByVal DateLastUsed As Object = Nothing)
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True

            cmbTankStatus.DataSource = pTank.PopulateTankStatus(Mode, DateLastUsed)
            cmbTankStatus.DisplayMember = "PROPERTY_NAME"
            cmbTankStatus.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tank Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub PopulateTankManufacturer()
        Try
            cmbTankManufacturer.DataSource = pTank.PopulateTankManufacturer
            cmbTankManufacturer.DisplayMember = "PROPERTY_NAME"
            cmbTankManufacturer.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateTankMaterial()
        Try
            cmbTankMaterial.DataSource = pTank.PopulateTankMaterialOfConstruction
            cmbTankMaterial.DisplayMember = "PROPERTY_NAME"
            cmbTankMaterial.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateTankOptions(Optional ByVal tnkMatDesc As Integer = 0)
        Try
            'added by Hua Cao 09/17/2008    If tank install date is greater or equal to 10/01/08, then Secondary Tank Option is filtered for like ‘%Double Walled%’.
            If Me.dtPickTankInstalled.Value >= Convert.ToDateTime("10-01-2008") Then
                Dim strSQL As String
                Dim dtReturn As DataTable
                If tnkMatDesc <> 0 Then
                    strSQL = "vSECONDARYTANKOPTIONSTYPE where property_name like '%Double-Walled%' and PROPERTY_ID_PARENT = " + tnkMatDesc.ToString
                Else
                    strSQL = "vSECONDARYTANKOPTIONSTYPE where property_name like '%Double-Walled%'"
                End If
                dtReturn = pTank.GetDataTable(strSQL)
                Me.cmbTankOptions.DataSource = dtReturn
            Else
                cmbTankOptions.DataSource = pTank.PopulateTankSecondaryOption(tnkMatDesc)
            End If
            cmbTankOptions.DisplayMember = "PROPERTY_NAME"
            cmbTankOptions.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateTankCPType()
        Try
            cmbTankCPType.DataSource = pTank.PopulateTankCPType
            cmbTankCPType.DisplayMember = "PROPERTY_NAME"
            cmbTankCPType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateTankOverfillProtectionType()
        Try
            cmbTankOverfillProtectionType.DataSource = pTank.PopulateTankOverFillProtectionType
            cmbTankOverfillProtectionType.DisplayMember = "PROPERTY_NAME"
            cmbTankOverfillProtectionType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateTankReleaseDetection(Optional ByVal tnkModDesc As Integer = 0, Optional ByVal tankCapacity As Integer = 0)
        Try
            '   Added by Hua Cao 09/22/2008 If tank install date is greater or equal to 10/01/08, then Release Detection is filtered for like ‘%Interstitial Monitoring%’ 
            If Me.dtPickTankInstalled.Value >= Convert.ToDateTime("10-01-2008") Then
                Dim strSQL As String
                Dim dtReturn As DataTable
                If tnkModDesc <> 0 Then
                    strSQL = "VRELEASEDETECTIONTYPE where property_name like '%Interstitial Monitoring%' and PROPERTY_ID_PARENT = " + tnkModDesc.ToString
                Else
                    strSQL = "VRELEASEDETECTIONTYPE where property_name like '%Interstitial Monitoring%'"
                End If
                dtReturn = pTank.GetDataTable(strSQL)

                Me.cmbTankReleaseDetection.DataSource = dtReturn
            Else
                cmbTankReleaseDetection.DataSource = pTank.PopulateTankReleaseDetection(tnkModDesc, tankCapacity)
            End If
            cmbTankReleaseDetection.DisplayMember = "PROPERTY_NAME"
            cmbTankReleaseDetection.ValueMember = "PROPERTY_ID"
            cmbTankReleaseDetection.SelectedValue = -1
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateTankClosureType()
        Try
            cmbTankClosureType.DataSource = pTank.PopulateClosureType
            cmbTankClosureType.DisplayMember = "PROPERTY_NAME"
            cmbTankClosureType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateTankInertFill()
        Try
            cmbTankInertFill.DataSource = pTank.PopulateTankPipeInertFill
            cmbTankInertFill.DisplayMember = "PROPERTY_NAME"
            cmbTankInertFill.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub CheckTankStatus(ByVal oOldStatus As Integer, ByVal nNewStatus As Integer)
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            '424 - CIU
            '425 - TOS
            '426 - POU
            '428 - RegPending
            '429 - TOSI
            '430 - Unregulated
            bolLoading = True

            If oOldStatus = 424 And (nNewStatus = 425 Or nNewStatus = 429) Then ' CIU TO TOS OR TOSI
                ' set focus to date last used
                dtPickLastUsed.Focus()
            ElseIf oOldStatus = 425 Or oOldStatus = 429 And nNewStatus = 424 Then ' TOS OR TOSI TO CIU
                dtPickDatePlacedInService.Focus()
            End If
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub

    Private Sub SetupCompartmentsTable()
        Try
            dTableTankCompartments = New DataTable

            dTableTankCompartments.Columns.Add("ID", GetType(String))
            dTableTankCompartments.Columns.Add("TANK_ID", GetType(Integer))
            dTableTankCompartments.Columns.Add("COMPARTMENT NUMBER", GetType(Integer))
            dTableTankCompartments.Columns.Add("COMPARTMENT #", GetType(Integer))
            dTableTankCompartments.Columns.Add("CAPACITY", GetType(Integer))
            dTableTankCompartments.Columns("CAPACITY").DefaultValue = 0
            dTableTankCompartments.Columns.Add("SUBSTANCE", GetType(Integer))
            dTableTankCompartments.Columns.Add("CERCLA#", GetType(Integer))
            dTableTankCompartments.Columns.Add("FUEL TYPE ID", GetType(Integer))
            dTableTankCompartments.Columns.Add("MANIFOLD INFO", GetType(String))
            dTableTankCompartments.Columns("MANIFOLD INFO").DefaultValue = "N/A"
            dTableTankCompartments.Columns.Add("Deleted", GetType(Boolean))
            dTableTankCompartments.Columns.Add("Created By", GetType(String))
            dTableTankCompartments.Columns.Add("Created On", GetType(DateTime))
            dTableTankCompartments.Columns.Add("Modified By", GetType(String))
            dTableTankCompartments.Columns.Add("Modified On", GetType(DateTime))
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateCompartment(ByVal compNum As Integer)
        Try
            ClearCompartmentArea()
            ResetCompartmentNumber()
            nCompartmentNumber = compNum

            txtNonCompTankCapacity.Text = pTank.Compartments.Capacity.ToString

            PopulateCompartmentSubstance()
            UIUtilsGen.SetComboboxItemByValue(cmbTanksubstance, pTank.Compartments.Substance)
            If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTanksubstance, pTank.Compartments.Substance, "PROPERTY_ID", "=") Then
                pTank.Compartments.Substance = 0
            End If
            If pTank.Compartments.Substance <= 0 Then
                cmbTanksubstance.SelectedIndex = -1
                If cmbTanksubstance.SelectedIndex <> -1 Then
                    cmbTanksubstance.SelectedIndex = -1
                End If
            End If

            PopulateCompartmentFuelType(pTank.Compartments.Substance)
            ' if no values returned from property relation table, disble field
            If cmbTankFuelType.Items.Count = 0 Then
                cmbTankFuelType.Enabled = False
                pTank.Compartments.FuelTypeId = 0
            Else
                cmbTankFuelType.Enabled = True
                UIUtilsGen.SetComboboxItemByValue(cmbTankFuelType, pTank.Compartments.FuelTypeId)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankFuelType, pTank.Compartments.FuelTypeId, "PROPERTY_ID", "=") Then
                    pTank.Compartments.FuelTypeId = 0
                End If
            End If
            If pTank.Compartments.FuelTypeId <= 0 Then
                cmbTankFuelType.SelectedIndex = -1
                If cmbTankFuelType.SelectedIndex <> -1 Then
                    cmbTankFuelType.SelectedIndex = -1
                End If
            End If

            If cmbTanksubstance.Text.IndexOf("Hazardous Substance") > -1 Then
                cmbTankCercla.Enabled = True
                cmbTankCerclaDesc.Enabled = True
                PopulateCompartmentCercla()
                PopulateCompartmentCerclaDesc()
                UIUtilsGen.SetComboboxItemByValue(cmbTankCercla, pTank.Compartments.CCERCLA)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankCercla, pTank.Compartments.CCERCLA, "PROPERTY_ID", "=") Then
                    pTank.Compartments.CCERCLA = 0
                End If
                UIUtilsGen.SetComboboxItemByValue(cmbTankCerclaDesc, pTank.Compartments.CCERCLA)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankCerclaDesc, pTank.Compartments.CCERCLA, "PROPERTY_ID", "=") Then
                    pTank.Compartments.CCERCLA = 0
                End If
                If pTank.Compartments.CCERCLA <= 0 Then
                    cmbTankCercla.SelectedIndex = -1
                    cmbTankCerclaDesc.SelectedIndex = -1
                    If cmbTankCercla.SelectedIndex <> -1 Then
                        cmbTankCercla.SelectedIndex = -1
                    End If
                    If cmbTankCerclaDesc.SelectedIndex <> -1 Then
                        cmbTankCerclaDesc.SelectedIndex = -1
                    End If
                End If
            Else
                pTank.Compartments.CCERCLA = 0
                cmbTankCercla.Enabled = False
                cmbTankCercla.DataSource = Nothing
                cmbTankCerclaDesc.Enabled = False
                cmbTankCerclaDesc.DataSource = Nothing
            End If

            ' Manifold
            Dim dtManifold As DataTable = pTank.Compartments.getManifold(nTankID)
            Dim drRows() As DataRow
            Dim strManifold As String = ""

            If Not dtManifold Is Nothing Then
                For Each drManifold As DataRow In dtManifold.Rows
                    If Not drManifold("Manifold") Is DBNull.Value Then
                        strManifold += drManifold("Manifold").ToString + ", "
                    End If
                Next
                If strManifold <> String.Empty Then
                    strManifold = strManifold.Trim.TrimEnd(",")
                Else
                    strManifold = "N/A"
                End If
            Else
                strManifold = "N/A"
            End If
            lblTankManifoldValue.Text = strManifold
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateCompartments(ByVal tnkID As Integer)
        Try
            ClearCompartmentArea()
            ResetCompartmentNumber()

            pnlNonCompProperties.Visible = False
            txtNonCompTankCapacity.Text = String.Empty
            cmbTanksubstance.DataSource = Nothing
            cmbTankFuelType.DataSource = Nothing
            cmbTankCercla.DataSource = Nothing
            cmbTankCerclaDesc.DataSource = Nothing

            txtNonCompTankCapacity.Enabled = False
            cmbTanksubstance.Enabled = False
            cmbTankFuelType.Enabled = False
            cmbTankCercla.Enabled = False
            cmbTankCerclaDesc.Enabled = False

            chkTankCompartment.Checked = True

            txtTankCapacity.Enabled = True
            txtTankCompartmentNumber.Enabled = True
            dGridCompartments.Enabled = True

            lblTankCapacity.Visible = True
            txtTankCapacity.Visible = True
            lblTankCompartmentNumber.Visible = True
            txtTankCompartmentNumber.Visible = True
            dGridCompartments.Visible = True

            SetupCompartmentsTable()
            dTableTankCompartments = pTank.Compartments.EntityTable
            'Dim dtComps As DataTable = pTank.Compartments.EntityTable
            Dim dtManifold As DataTable = pTank.Compartments.getManifold(nTankID)
            Dim drRows() As DataRow

            'Dim compCapacity As Integer = 0
            'Dim compCount As Integer = 0

            For Each dr As DataRow In dTableTankCompartments.Rows
                If Not dtManifold Is Nothing Then
                    drRows = dtManifold.Select("compartment_number=" + CType(dr.Item("COMPARTMENT NUMBER"), String))
                    For Each drManifold As DataRow In drRows
                        If drManifold("Compartment_Number") = dr("COMPARTMENT NUMBER") Then
                            If Not drManifold("Manifold") Is DBNull.Value Then
                                dr("MANIFOLD INFO") = drManifold("Manifold")
                            End If
                        End If
                    Next
                End If
                'compCount += 1
                'compCapacity += IIf(dr("CAPACITY") Is DBNull.Value, 0, dr("CAPACITY"))
            Next

            dGridCompartments.DataSource = dTableTankCompartments
            dGridCompartments.DrawFilter = rp

            'txtTankCapacity.Text = compCapacity.ToString
            'txtTankCompartmentNumber.Text = compCount.ToString
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub PopulateCompartmentSubstance()
        Try
            cmbTanksubstance.DataSource = pTank.PopulateCompartmentSubstance
            cmbTanksubstance.DisplayMember = "PROPERTY_NAME"
            cmbTanksubstance.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateCompartmentFuelType(Optional ByVal compSubstance As Int64 = 0)
        Try
            cmbTankFuelType.DataSource = pTank.PopulateCompartmentFuelTypes(compSubstance)
            cmbTankFuelType.DisplayMember = "PROPERTY_NAME"
            cmbTankFuelType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateCompartmentCercla()
        Try
            cmbTankCercla.DataSource = pTank.PopulateCERCLA
            cmbTankCercla.DisplayMember = "PROPERTY_NAME"
            cmbTankCercla.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateCompartmentCerclaDesc()
        Try
            cmbTankCerclaDesc.DataSource = pTank.PopulateCERCLA
            cmbTankCerclaDesc.DisplayMember = "PROPERTY_DESC"
            cmbTankCerclaDesc.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Function GetPrevNextTank(ByVal tnkID As Integer, ByVal getNext As Boolean) As Integer
        Try
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim retVal As Integer = tnkID
            Dim sl As New SortedList
            Dim slRow As New SortedList

            For Each ugRow In dgPipesAndTanks.Rows
                sl.Add(ugRow.Cells("TANK ID").Value, ugRow.Cells("TANK ID").Value)
                slRow.Add(ugRow.Cells("TANK ID").Value, ugRow)
            Next

            retVal = GetPrevNext(sl, getNext, tnkID)
            ugTankRow = slRow.Item(retVal)
            Return retVal
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Function

    Private Sub cmbTankStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTankStatus.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.TankStatus = UIUtilsGen.GetComboBoxValueInt(cmbTankStatus)
            If pTank.TankStatusOriginal > 0 Then
                FillTankFormFields(True, False)
            Else
                FillTankFormFields(True, True)
            End If
            CheckTankStatus(pTank.TankStatusOriginal, pTank.TankStatus)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickTankInstalled_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickTankInstalled.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickTankInstalled)
            pTank.DateInstalledTank = UIUtilsGen.GetDatePickerValue(dtPickTankInstalled)
            'Added by Hua Cao 09/22/2008   If tank install date is greater or equal to 10/01/08, then
            '-- Release Detection is filtered for like ‘%Interstitial Monitoring%’ 
            '-- Secondary Tank Option is filtered for like ‘%Double Walled%’.
            Me.PopulateTankReleaseDetection(pTank.TankModDesc)
            Me.PopulateTankOptions(pTank.TankMatDesc)
        Catch ex As Exception
            ShowError(ex)
        End Try

    End Sub
    Private Sub dtPickDatePlacedInService_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickDatePlacedInService.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickDatePlacedInService)
            pTank.PlacedInServiceDate = UIUtilsGen.GetDatePickerValue(dtPickDatePlacedInService)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbTankManufacturer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTankManufacturer.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.TankManufacturer = UIUtilsGen.GetComboBoxValueInt(cmbTankManufacturer)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkTankCompartment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTankCompartment.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            bolLoading = True
            If chkTankCompartment.Checked Then
                pTank.Compartment = chkTankCompartment.Checked
                ' clearing all the compartments from collection
                'pTank.TankInfo.CompartmentCollection = New MUSTER.Info.CompartmentCollection
                'pTank.Compartments = New MUSTER.BusinessLogic.pCompartment
                pTank.Compartments.Retrieve(pTank.TankInfo, nTankID, False)

                PopulateCompartments(nTankID)
            Else
                If dGridCompartments.Rows.Count > 0 Then
                    If MsgBox("Changing a Compartmentalized tank to a Non-Compartmentalized tank will result in loss of Data in the Grid Below" + vbCrLf + vbCrLf + _
                                "Do You want to continue?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Try
                            bolLoading = True
                            chkTankCompartment.Checked = Not chkTankCompartment.Checked
                            Exit Sub
                        Catch ex As Exception
                            Throw ex
                        Finally
                            bolLoading = False
                        End Try
                    End If
                End If

                pTank.Compartment = chkTankCompartment.Checked
                pTank.Compartments.TankId = pTank.TankId

                Dim compNumToUse As Integer = 0

                If dGridCompartments.Rows.Count > 0 Then
                    compNumToUse = dGridCompartments.Rows(0).Cells("COMPARTMENT NUMBER").Value
                End If

                For Each comp As MUSTER.Info.CompartmentInfo In pTank.TankInfo.CompartmentCollection.Values
                    If comp.COMPARTMENTNumber <> compNumToUse Then
                        comp.Deleted = True
                    End If
                Next

                pTank.Compartments.Retrieve(pTank.TankInfo, nTankID.ToString + "|" + compNumToUse.ToString, False)

                PopulateCompartment(compNumToUse)

            End If
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = False
        End Try
    End Sub

    'dGridCompartments events
    Private Sub getSummaryValue(ByVal uGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByRef lblCapacity As Label, ByRef lblCount As Label)
        If uGrid.Rows.SummaryValues.Count > 0 Then
            lblCapacity.Text = uGrid.Rows.SummaryValues.Item(0).Value.ToString
            lblCount.Text = uGrid.Rows.SummaryValues.Item(1).Value.ToString
        Else
            lblCapacity.Text = 0
            lblCount.Text = 0
        End If
    End Sub
    Private Sub SetugRowComboValue(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Dim vListFuelType As Infragistics.Win.ValueList
        Try
            ' Substance
            If vListSubstance.FindByDataValue(ug.Cells("SUBSTANCE").Value) Is Nothing Then
                ug.Cells("SUBSTANCE").Value = DBNull.Value
                ug.Cells("FUEL TYPE ID").Value = DBNull.Value
                ug.Cells("FUEL TYPE ID").Hidden = True
            Else
                vListFuelType = New Infragistics.Win.ValueList
                Dim dt As DataTable = pTank.PopulateCompartmentFuelTypes(ug.Cells("SUBSTANCE").Value)
                If dt Is Nothing Then
                    ug.Cells("FUEL TYPE ID").Value = DBNull.Value
                    ug.Cells("FUEL TYPE ID").Hidden = True
                Else
                    If dt.Rows.Count > 0 Then
                        ug.Cells("FUEL TYPE ID").Hidden = False
                        For Each row As DataRow In dt.Rows
                            vListFuelType.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ug.Cells("FUEL TYPE ID").ValueList = vListFuelType
                        If vListFuelType.FindByDataValue(ug.Cells("FUEL TYPE ID").Value) Is Nothing Then
                            ug.Cells("FUEL TYPE ID").Value = DBNull.Value
                        End If
                    Else
                        ug.Cells("FUEL TYPE ID").Value = DBNull.Value
                        ug.Cells("FUEL TYPE ID").Hidden = True
                    End If
                End If
            End If

            If ug.Cells("SUBSTANCE").Text.IndexOf("Hazardous Substance") > -1 Then
                ug.Cells("CERCLA#").Hidden = False
                ' Cercla Number
                If vListCerclaNumber.FindByDataValue(ug.Cells("CERCLA#").Value) Is Nothing Then
                    ug.Cells("CERCLA#").Value = DBNull.Value
                End If
            Else
                ug.Cells("CERCLA#").Value = DBNull.Value
                ug.Cells("CERCLA#").Hidden = True
            End If

            ' Fuel Type
            'If vListFuelType.FindByDataValue(ug.Cells("FUEL TYPE ID").Value) Is Nothing Then
            '    ug.Cells("FUEL TYPE ID").Value = DBNull.Value
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub dGridCompartments_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles dGridCompartments.InitializeLayout
        Try
            e.Layout.Bands(0).Columns("COMPARTMENT #").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            If e.Layout.Bands(0).Columns.Exists("CAPACITY") Then
                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("CAPACITY"), Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn)
            End If
            If e.Layout.Bands(0).Columns.Exists("COMPARTMENT #") Then
                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Count, e.Layout.Bands(0).Columns("COMPARTMENT #"), Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn)
            End If
            e.Layout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.False

            e.Layout.Bands(0).Columns("CAPACITY").MaskInput = "nnnnnnnnn"
            e.Layout.Bands(0).Columns("CAPACITY").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(0).Columns("CAPACITY").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            e.Layout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom
            e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True

            e.Layout.Bands(0).Columns("ID").Hidden = True
            e.Layout.Bands(0).Columns("TANK_ID").Hidden = True
            e.Layout.Bands(0).Columns("COMPARTMENT NUMBER").Hidden = True
            e.Layout.Bands(0).Columns("Deleted").Hidden = True
            e.Layout.Bands(0).Columns("Created By").Hidden = True
            e.Layout.Bands(0).Columns("Created On").Hidden = True
            e.Layout.Bands(0).Columns("Modified By").Hidden = True
            e.Layout.Bands(0).Columns("Modified On").Hidden = True

            e.Layout.Bands(0).Columns("COMPARTMENT #").Width = 90
            e.Layout.Bands(0).Columns("CAPACITY").Width = 100
            e.Layout.Bands(0).Columns("SUBSTANCE").Width = 150
            e.Layout.Bands(0).Columns("CERCLA#").Width = 100
            e.Layout.Bands(0).Columns("FUEL TYPE ID").Width = 150
            e.Layout.Bands(0).Columns("MANIFOLD INFO").Width = 300

            e.Layout.Bands(0).Columns("SUBSTANCE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown
            e.Layout.Bands(0).Columns("CERCLA#").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown
            e.Layout.Bands(0).Columns("FUEL TYPE ID").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown

            e.Layout.Bands(0).Columns("COMPARTMENT #").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("MANIFOLD INFO").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(0).Columns("COMPARTMENT #").Header.Caption = "Compartment #"
            e.Layout.Bands(0).Columns("CAPACITY").Header.Caption = "Capacity"
            e.Layout.Bands(0).Columns("SUBSTANCE").Header.Caption = "Substance"
            e.Layout.Bands(0).Columns("CERCLA#").Header.Caption = "Cercla Number"
            e.Layout.Bands(0).Columns("FUEL TYPE ID").Header.Caption = "Fuel Type"
            e.Layout.Bands(0).Columns("MANIFOLD INFO").Header.Caption = "Manifold Info"

            If e.Layout.Bands(0).Columns("SUBSTANCE").ValueList Is Nothing Then
                vListSubstance = New Infragistics.Win.ValueList
                For Each row As DataRow In pTank.PopulateTankSubstance.Rows
                    vListSubstance.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(0).Columns("SUBSTANCE").ValueList = vListSubstance
            End If

            If e.Layout.Bands(0).Columns("CERCLA#").ValueList Is Nothing Then
                vListCerclaNumber = New Infragistics.Win.ValueList
                For Each row As DataRow In pTank.PopulateCERCLA.Rows
                    vListCerclaNumber.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(0).Columns("CERCLA#").ValueList = vListCerclaNumber
            End If

            'If e.Layout.Bands(0).Columns("FUEL TYPE ID").ValueList Is Nothing Then
            '    vListFuelType = New Infragistics.Win.ValueList
            '    For Each row As DataRow In pTank.PopulateCompartmentFuelTypes.Rows
            '        vListFuelType.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
            '    Next
            '    e.Layout.Bands(0).Columns("FUEL TYPE ID").ValueList = vListFuelType
            'End If

            Dim maxCompNum As Integer = 0
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Layout.Grid.Rows
                'ugRow.Cells("SUBSTANCE").ValueList = vListSubstance
                'ugRow.Cells("CERCLA#").ValueList = vListCerclaNumber
                'ugRow.Cells("FUEL TYPE ID").ValueList = vListFuelType
                SetugRowComboValue(ugRow)

                If Not ugRow.Cells("COMPARTMENT #").Value Is DBNull.Value Then
                    If ugRow.Cells("COMPARTMENT #").Value > maxCompNum Then
                        maxCompNum = ugRow.Cells("COMPARTMENT #").Value
                    End If
                End If
            Next

            If maxCompNum <= 0 Then
                maxCompNum = 0
                e.Layout.Grid.Rows(0).Cells("COMPARTMENT #").Value = maxCompNum + 1
            End If

            getSummaryValue(sender, txtTankCapacity, txtTankCompartmentNumber)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub dGridCompartments_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles dGridCompartments.CellChange
        Try
            If "CAPACITY".Equals(e.Cell.Column.Key) Then
                If bolAddComp Then bolAddComp = False
                If e.Cell.EditorResolved.Value Is DBNull.Value Then
                    e.Cell.Value = 0
                Else
                    e.Cell.Value = e.Cell.EditorResolved.Value
                End If
                If pTank.Compartments.COMPARTMENTNumber <> e.Cell.Row.Cells("COMPARTMENT NUMBER").Value Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, nTankID.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT NUMBER").Value.ToString, False)
                End If
                pTank.Compartments.Capacity = e.Cell.Value
                getSummaryValue(dGridCompartments, txtTankCapacity, txtTankCompartmentNumber)
            ElseIf "SUBSTANCE".Equals(e.Cell.Column.Key) Then
                If bolAddComp Then bolAddComp = False
                If pTank.Compartments.COMPARTMENTNumber <> e.Cell.Row.Cells("COMPARTMENT NUMBER").Value Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, nTankID.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT NUMBER").Value.ToString, False)
                End If
                pTank.Compartments.Substance = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                e.Cell.Value = pTank.Compartments.Substance
                SetugRowComboValue(e.Cell.Row)
            ElseIf "CERCLA#".Equals(e.Cell.Column.Key) Then
                If bolAddComp Then bolAddComp = False
                If pTank.Compartments.COMPARTMENTNumber <> e.Cell.Row.Cells("COMPARTMENT NUMBER").Value Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, nTankID.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT NUMBER").Value.ToString, False)
                End If
                pTank.Compartments.CCERCLA = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                e.Cell.Value = pTank.Compartments.CCERCLA
            ElseIf "FUEL TYPE ID".Equals(e.Cell.Column.Key) Then
                If bolAddComp Then bolAddComp = False
                If pTank.Compartments.COMPARTMENTNumber <> e.Cell.Row.Cells("COMPARTMENT NUMBER").Value Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, nTankID.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT NUMBER").Value.ToString, False)
                End If
                pTank.Compartments.FuelTypeId = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                e.Cell.Value = pTank.Compartments.FuelTypeId
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dGridCompartments_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles dGridCompartments.AfterCellUpdate
    '    If bolLoading Then Exit Sub
    '    Try
    '        If "CAPACITY".Equals(e.Cell.Column.Key) Then
    '            FillTankFormFields(True, False)
    '        End If
    '    Catch ex As Exception
    '        ShowError(ex)
    '    End Try
    'End Sub
    Private Sub dGridCompartments_BeforeRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs) Handles dGridCompartments.BeforeRowInsert
        Try
            If dGridCompartments.Rows.Count > 0 Then
                pTank.Compartments.Retrieve(pTank.TankInfo, pTank.TankId.ToString + "|0")
                pTank.Compartments.ChangeCompartmentNumberKey(, pTank.TankId)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dGridCompartments_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dGridCompartments.AfterRowInsert
        Try
            e.Row.Cells("ID").Value = pTank.Compartments.ID
            e.Row.Cells("COMPARTMENT NUMBER").Value = pTank.Compartments.COMPARTMENTNumber
            e.Row.Cells("CERCLA#").Hidden = True
            e.Row.Cells("MANIFOLD INFO").Value = "-NA-"
            getSummaryValue(dGridCompartments, txtTankCapacity, txtTankCompartmentNumber)

            Dim maxCompNum As Integer = 0
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In dGridCompartments.Rows
                If Not ugRow.Cells("COMPARTMENT #").Value Is DBNull.Value Then
                    If ugRow.Cells("COMPARTMENT #").Value > maxCompNum Then
                        maxCompNum = ugRow.Cells("COMPARTMENT #").Value
                    End If
                End If
            Next
            If maxCompNum < 0 Then maxCompNum = 0
            dGridCompartments.ActiveRow.Cells("COMPARTMENT #").Value = maxCompNum + 1

            dGridCompartments.ActiveCell = e.Row.Cells("CAPACITY")
            dGridCompartments.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ToggleCellSel, False, False)
            dGridCompartments.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    ' do we need this
    Private Sub dGridCompartments_BeforeCellActivated(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles dGridCompartments.BeforeCellActivate
        Try
            If "CERCLA#".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("Substance").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToUpper.IndexOf("HAZARDOUS") < 0 Then
                    e.Cell.Value = Nothing
                    e.Cancel = True
                End If
            ElseIf "MANIFOLD INFO".Equals(e.Cell.Column.Key) Or "COMPARTMENT #".Equals(e.Cell.Column.Key) Then
                e.Cancel = True
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dGridCompartments_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles dGridCompartments.AfterCellUpdate
    '    Try
    '        fillCompartment()
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub txtTankCapacity_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTankCapacity.TextChanged
        If bolLoading Then Exit Sub
        Try
            If chkTankCompartment.Checked Then
                FillTankFormFields(True, False)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub dGridCompartments_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles dGridCompartments.BeforeRowsDeleted
        Try
            e.DisplayPromptMsg = False
            If Not dGridCompartments.ActiveRow Is Nothing Then
                pTank.Compartments.Retrieve(pTank.TankInfo, pTank.TankId.ToString + "|" + dGridCompartments.ActiveRow.Cells("COMPARTMENT NUMBER").Value.ToString, False)
                If pTank.Compartments.Capacity <> 0 Or pTank.Compartments.Substance <> 0 Or pTank.Compartments.FuelTypeId = 0 Or pTank.Compartments.CCERCLA = 0 Then
                    If MessageBox.Show("Are you sure you wish to DELETE the Record?", "MUSTER", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                        Exit Sub
                    End If
                End If
                Dim compHasPipes As Boolean = False
                If Not ugTankRow.ChildBands Is Nothing Then
                    For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugTankRow.ChildBands(0).Rows
                        If Not ugRow.Cells("COMPARTMENT NUMBER").Value Is DBNull.Value Then
                            If ugRow.Cells("COMPARTMENT NUMBER").Value = dGridCompartments.ActiveRow.Cells("COMPARTMENT NUMBER").Value Then
                                compHasPipes = True
                                Exit For
                            End If
                        End If
                    Next
                End If
                If compHasPipes Then
                    MsgBox("The Specified compartment has associated Pipe(s). Delete Pipe(s) before deleting the compartment")
                    e.Cancel = True
                Else
                    pTank.Compartments.Deleted = True
                    If pTank.Compartments.COMPARTMENTNumber <= 0 Then
                        pTank.Compartments.CreatedBy = MC.AppUser.ID
                    Else
                        pTank.Compartments.ModifiedBy = MC.AppUser.ID
                    End If
                    returnVal = String.Empty
                    pTank.Compartments.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID, True, True)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If


                    'Dim compCapacity As Integer = 0
                    'Dim compCount As Integer = 0
                    'If txtTankCapacity.Text <> String.Empty Then
                    '    compCapacity = CType(txtTankCapacity.Text, Integer)
                    '    compCapacity -= pTank.Compartments.Capacity
                    'End If

                    'If txtTankCompartmentNumber.Text <> String.Empty Then
                    '    compCount = CType(txtTankCompartmentNumber.Text, Integer)
                    '    compCount -= 1
                    'End If

                    'txtTankCapacity.Text = compCapacity.ToString
                    'txtTankCompartmentNumber.Text = compCount.ToString

                    pTank.Compartments.Remove(nTankID.ToString + "|" + pTank.Compartments.COMPARTMENTNumber.ToString)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error During Compartment Delete: " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dGridCompartments_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dGridCompartments.AfterRowsDeleted
        getSummaryValue(dGridCompartments, txtTankCapacity, txtTankCompartmentNumber)
    End Sub
    'Private Sub dGridCompartments_AfterCellActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles dGridCompartments.AfterCellActivate
    '    Try
    '        If Not dGridCompartments.ActiveRow Is Nothing Then
    '            If dGridCompartments.ActiveRow.Cells("COMPARTMENT #").Value Is DBNull.Value Then
    '            Else
    '                ' these 2 lines so it doesn't crash
    '                maxCompNum = dGridCompartments.ActiveRow.Cells("COMPARTMENT #").Value
    '                dGridCompartments.ActiveRow.Cells("COMPARTMENT #").Value = maxCompNum
    '            End If
    '        End If
    '    Catch ex As Exception
    '        ShowError(ex)
    '    End Try
    'End Sub
    'Private Sub dGridCompartments_BeforeRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dGridCompartments.BeforeRowActivate
    '    If Not e.Row.Cells("COMPARTMENT NUMBER").Value Is DBNull.Value Then
    '        If nCompartmentNumber <> e.Row.Cells("COMPARTMENT NUMBER").Value Then
    '            'If nCompartmentNumber < 0 And e.Row.Cells("COMPARTMENT NUMBER").Value > 0 Then
    '                'If bolAddComp Then
    '                'pTank.Compartments.Remove(nTankID.ToString + "|" + nCompartmentNumber.ToString)
    '                'End If
    '            'End If
    '            nCompartmentNumber = e.Row.Cells("COMPARTMENT NUMBER").Value
    '            'getSummaryValue(dGridCompartments, txtTankCapacity, txtTankCompartmentNumber)
    '        End If
    '    End If
    'End Sub
    'Private Sub dGridCompartments_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dGridCompartments.AfterRowUpdate
    'End Sub

    Private Sub txtNonCompTankCapacity_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNonCompTankCapacity.TextChanged
        If bolLoading Then Exit Sub
        Try
            If Not chkTankCompartment.Checked Then
                If txtNonCompTankCapacity.Text = String.Empty Then
                    pTank.Compartments.Capacity = 0
                Else
                    pTank.Compartments.Capacity = Integer.Parse(txtNonCompTankCapacity.Text)
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtNonCompTankCapacity_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNonCompTankCapacity.Leave
        If bolLoading Then Exit Sub
        Try
            FillTankFormFields(True, True)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbTanksubstance_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTanksubstance.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            If Not chkTankCompartment.Checked Then

                pTank.Compartments.Substance = UIUtilsGen.GetComboBoxValueInt(cmbTanksubstance)

                PopulateCompartmentFuelType(pTank.Compartments.Substance)
                ' if no values returned from property relation table, disble field
                If cmbTankFuelType.Items.Count = 0 Then
                    cmbTankFuelType.Enabled = False
                    pTank.Compartments.FuelTypeId = 0
                Else
                    cmbTankFuelType.Enabled = True
                    UIUtilsGen.SetComboboxItemByValue(cmbTankFuelType, pTank.Compartments.FuelTypeId)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankFuelType, pTank.Compartments.FuelTypeId, "PROPERTY_ID", "=") Then
                        pTank.Compartments.FuelTypeId = 0
                    End If
                End If
                If pTank.Compartments.FuelTypeId <= 0 Then
                    cmbTankFuelType.SelectedIndex = -1
                    If cmbTankFuelType.SelectedIndex <> -1 Then
                        cmbTankFuelType.SelectedIndex = -1
                    End If
                End If

                If cmbTanksubstance.Text.IndexOf("Hazardous Substance") > -1 Then
                    cmbTankCercla.Enabled = True
                    cmbTankCerclaDesc.Enabled = True
                    PopulateCompartmentCercla()
                    PopulateCompartmentCerclaDesc()
                    cmbTankCercla.SelectedIndex = -1
                    If cmbTankCercla.SelectedIndex <> -1 Then
                        cmbTankCercla.SelectedIndex = -1
                    End If
                    cmbTankCerclaDesc.SelectedIndex = -1
                    If cmbTankCerclaDesc.SelectedIndex <> -1 Then
                        cmbTankCerclaDesc.SelectedIndex = -1
                    End If
                Else
                    pTank.Compartments.CCERCLA = 0
                    cmbTankCercla.Enabled = False
                    cmbTankCercla.DataSource = Nothing
                    cmbTankCerclaDesc.Enabled = False
                    cmbTankCerclaDesc.DataSource = Nothing
                End If
                If cmbTanksubstance.Text.IndexOf("Used Oil") > -1 Then
                    cmbTankOverfillProtectionType.SelectedIndex = -1
                    If cmbTankOverfillProtectionType.SelectedIndex <> -1 Then
                        cmbTankOverfillProtectionType.SelectedIndex = -1
                    End If
                    pTank.OverFillType = 0
                    cmbTankOverfillProtectionType.Enabled = False
                Else
                    If chkDeliveriesLimited.Checked Then
                        cmbTankOverfillProtectionType.SelectedIndex = -1
                        If cmbTankOverfillProtectionType.SelectedIndex <> -1 Then
                            cmbTankOverfillProtectionType.SelectedIndex = -1
                        End If
                        pTank.OverFillType = 0
                        cmbTankOverfillProtectionType.Enabled = False
                    Else
                        cmbTankOverfillProtectionType.Enabled = True
                        PopulateTankOverfillProtectionType()
                        UIUtilsGen.SetComboboxItemByValue(cmbTankOverfillProtectionType, pTank.OverFillType)
                        If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbTankOverfillProtectionType, pTank.OverFillType, "PROPERTY_ID", "=") Then
                            pTank.OverFillType = 0
                        End If
                        If pTank.OverFillType <= 0 Then
                            cmbTankOverfillProtectionType.SelectedIndex = -1
                            If cmbTankOverfillProtectionType.SelectedIndex <> -1 Then
                                cmbTankOverfillProtectionType.SelectedIndex = -1
                            End If
                        End If
                    End If
                End If
            End If

            If Me.cmbTanksubstance.SelectedValue = 314 Then ' Used Oil
                UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionInstalled, dtNullDate)
                UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionLastTested, dtNullDate)
                UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionLastInspected, dtNullDate)
                UIUtilsGen.SetDatePickerValue(dtPickOverfillPreventionInstalled, dtNullDate)
                Me.dtPickSpillPreventionInstalled.Enabled = False
                Me.dtPickSpillPreventionLastTested.Enabled = False
                Me.dtPickOverfillPreventionLastInspected.Enabled = False
                Me.dtPickOverfillPreventionInstalled.Enabled = False
                '    Me.dtPickSpillPreventionInstalled.Value = "1900-01-01"
                '   Me.dtPickSpillPreventionLastTested.Value = "1900-01-01"
                '  Me.dtPickOverfillPreventionInstalled.Value = "1900-01-01"
                ' Me.dtPickOverfillPreventionLastInspected.Value = "1900-01-01"
            Else
                If Not (pTank.SmallDelivery) Then
                    Me.dtPickSpillPreventionInstalled.Enabled = True
                    Me.dtPickSpillPreventionLastTested.Enabled = True
                    Me.dtPickOverfillPreventionLastInspected.Enabled = True
                    Me.dtPickOverfillPreventionInstalled.Enabled = True
                Else
                    Me.dtPickSpillPreventionInstalled.Enabled = False
                    Me.dtPickSpillPreventionLastTested.Enabled = False
                    Me.dtPickOverfillPreventionLastInspected.Enabled = False
                    Me.dtPickOverfillPreventionInstalled.Enabled = False
                    ' Me.dtPickSpillPreventionInstalled.Value = "1900-01-01"
                    'Me.dtPickSpillPreventionLastTested.Value = "1900-01-01"
                    'Me.dtPickOverfillPreventionInstalled.Value = "1900-01-01"
                    'Me.dtPickOverfillPreventionLastInspected.Value = "1900-01-01"
                End If
            End If

        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub cmbTankFuelType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTankFuelType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            If Not Me.chkTankCompartment.Checked Then
                pTank.Compartments.FuelTypeId = UIUtilsGen.GetComboBoxValueInt(cmbTankFuelType)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbTankCercla_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTankCercla.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            If Not Me.chkTankCompartment.Checked Then
                If cmbTankCercla.SelectedValue Is Nothing Then
                    pTank.Compartments.CCERCLA = 0
                    cmbTankCerclaDesc.SelectedIndex = -1
                    If cmbTankCerclaDesc.SelectedIndex <> -1 Then
                        cmbTankCerclaDesc.SelectedIndex = -1
                    End If
                    ttCERCLA.AutoPopDelay = 15000 ' set lenth tip is visible to 15 seconds
                    ttCERCLA.SetToolTip(lblCERCLAtt, "None")
                Else
                    pTank.Compartments.CCERCLA = UIUtilsGen.GetComboBoxValueInt(cmbTankCercla)
                    cmbTankCerclaDesc.SelectedValue = cmbTankCercla.SelectedValue
                    ttCERCLA.AutoPopDelay = 15000 ' set lenth tip is visible to 15 seconds
                    If cmbTankCerclaDesc.Text.Trim = String.Empty Then
                        ttCERCLA.SetToolTip(lblCERCLAtt, "None")
                    Else
                        ttCERCLA.SetToolTip(lblCERCLAtt, cmbTankCerclaDesc.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmbTankMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTankMaterial.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.TankMatDesc = UIUtilsGen.GetComboBoxValueInt(cmbTankMaterial)
            FillTankFormFields(False, False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbTankOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTankOptions.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.TankModDesc = UIUtilsGen.GetComboBoxValueInt(cmbTankOptions)
            FillTankFormFields(False, False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbTankCPType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTankCPType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.TankCPType = UIUtilsGen.GetComboBoxValueInt(cmbTankCPType)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickCPInstalled_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickCPInstalled.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickCPInstalled)
            pTank.TCPInstallDate = UIUtilsGen.GetDatePickerValue(dtPickCPInstalled)
            FillTankFormFields(False, False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dtPickCPLastTested_ValueChanged()
    '    Dim dttemp, dtValidDate As Date
    '    Try
    '        dttemp = pTank.LastTCPDate
    '        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
    '        dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
    '        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
    '        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
    '            MsgBox("Tank CP Last Tested must be greater than or equal to " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub dtPickCPLastTested_ValueChangedHandler()
    '    Invoke(New HandlerDelegate(AddressOf dtPickCPLastTested_ValueChanged))
    'End Sub
    Private Sub dtPickCPLastTested_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickCPLastTested.CloseUp
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(dtPickCPLastTested)
            pTank.LastTCPDate = UIUtilsGen.GetDatePickerValue(dtPickCPLastTested)
            dtPickCPLastTested.Tag = pTank.LastTCPDate
            'Dim thread As New Threading.Thread(AddressOf dtPickCPLastTested_ValueChangedHandler)
            'thread.Start()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickCPLastTested_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickCPLastTested.TextChanged
        dtPickCPLastTested_ValueChanged(sender, e)
    End Sub
    Private Sub chkEmergencyPower_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkEmergencyPower.CheckedChanged
        If bolLoading Then Exit Sub
        Dim nTankCapacity As Integer
        Try
            pTank.TankEmergen = chkEmergencyPower.Checked
            FillTankFormFields(False, False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkDeliveriesLimited_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDeliveriesLimited.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            pTank.SmallDelivery = chkDeliveriesLimited.Checked
            FillTankFormFields(False, False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickInteriorLiningInstalled_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickInteriorLiningInstalled.CloseUp
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickInteriorLiningInstalled)
            pTank.LinedInteriorInstallDate = UIUtilsGen.GetDatePickerValue(dtPickInteriorLiningInstalled).Date
            dtPickInteriorLiningInstalled.Tag = pTank.LinedInteriorInstallDate
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickInteriorLiningInstalled_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickInteriorLiningInstalled.TextChanged
        dtPickInteriorLiningInstalled_ValueChanged(sender, e)
    End Sub
    'Private Sub dtPickLastInteriorLinningInspection_ValueChanged()
    '    Dim dttemp, dtValidDate As Date
    '    Try
    '        dttemp = pTank.LinedInteriorInspectDate
    '        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
    '        dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
    '        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
    '        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
    '            dtPickLastInteriorLinningInspection.Refresh()
    '            MsgBox("Last InteriorLinning Inspection must be greater than or equal to " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub dtPickLastInteriorLinningInspection_ValueChangedHandler()
    '    Invoke(New HandlerDelegate(AddressOf dtPickLastInteriorLinningInspection_ValueChanged))
    'End Sub
    Private Sub dtPickLastInteriorLinningInspection_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickLastInteriorLinningInspection.CloseUp
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickLastInteriorLinningInspection)
            pTank.LinedInteriorInspectDate = UIUtilsGen.GetDatePickerValue(dtPickLastInteriorLinningInspection)
            dtPickLastInteriorLinningInspection.Tag = pTank.LinedInteriorInspectDate
            'Dim thread As New Threading.Thread(AddressOf dtPickLastInteriorLinningInspection_ValueChangedHandler)
            'thread.Start()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickLastInteriorLinningInspection_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickLastInteriorLinningInspection.TextChanged
        dtPickLastInteriorLinningInspection_ValueChanged(sender, e)
    End Sub
    Private Sub cmbTankOverfillProtectionType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTankOverfillProtectionType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.OverFillType = cmbTankOverfillProtectionType.SelectedValue
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkBxSpillProtected_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkBxSpillProtected.Click
        If bolLoading Then Exit Sub
        Try
            pTank.SpillInstalled = chkBxSpillProtected.Checked
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkBxTightfillAdapters_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkBxTightfillAdapters.Click
        If bolLoading Then Exit Sub
        Try
            pTank.TightFillAdapters = chkBxTightfillAdapters.Checked
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOverFilledProtected_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOverFilledProtected.Click
        If bolLoading Then Exit Sub
        Try
            pTank.OverFillInstalled = chkOverFilledProtected.Checked
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmbTankReleaseDetection_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTankReleaseDetection.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.TankLD = UIUtilsGen.GetComboBoxValueInt(cmbTankReleaseDetection)
            If Not (Me.cmbTankReleaseDetection.DataSource Is Nothing) Then
                Me.cmbTankReleaseDetection.DisplayMember = "PROPERTY_NAME"
                Me.cmbTankReleaseDetection.ValueMember = "PROPERTY_ID"
            End If
            If Me.cmbTankReleaseDetection.SelectedValue = 336 Then 'Automatic Tank Gauging - 336
                Me.dtPickATGLastInspected.Enabled = True
            Else
                UIUtilsGen.SetDatePickerValue(dtPickATGLastInspected, dtNullDate)
                Me.dtPickATGLastInspected.Enabled = False
            End If
            If Me.cmbTankReleaseDetection.SelectedValue = 339 Then 'Electronic Interstitial Monitoring
                Me.dtPickElectronicDeviceInspected.Enabled = True
            Else
                UIUtilsGen.SetDatePickerValue(Me.dtPickElectronicDeviceInspected, dtNullDate)
                Me.dtPickElectronicDeviceInspected.Enabled = False
            End If
            If Me.cmbTankReleaseDetection.SelectedValue = 343 Then 'Visual Interstitial Monitoring
                Me.dtPickSecondaryContainmentLastInspected.Enabled = True
            Else
                UIUtilsGen.SetDatePickerValue(Me.dtPickSecondaryContainmentLastInspected, dtNullDate)
                Me.dtPickSecondaryContainmentLastInspected.Enabled = False
            End If
            FillTankFormFields(False, False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dtPickTankTightnessTest_ValueChanged()
    '    Dim dttemp, dtValidDate As Date
    '    Try
    '        dttemp = pTank.TTTDate
    '        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
    '        dtValidDate = DateAdd(DateInterval.Year, -5, dtValidDate)
    '        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
    '        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
    '            dtPickTankTightnessTest.Refresh()
    '            MsgBox("Tank Tightness Test must be greater than or equal to " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub dtPickTankTightnessTest_ValueChangedHandler()
    '    Invoke(New HandlerDelegate(AddressOf dtPickTankTightnessTest_ValueChanged))
    'End Sub
    Private Sub dtPickTankTightnessTest_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickTankTightnessTest.CloseUp
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickTankTightnessTest)
            pTank.TTTDate = UIUtilsGen.GetDatePickerValue(dtPickTankTightnessTest)
            dtPickTankTightnessTest.Tag = pTank.TTTDate
            'Dim thread As New Threading.Thread(AddressOf dtPickTankTightnessTest_ValueChangedHandler)
            'thread.Start()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickTankTightnessTest_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickTankTightnessTest.TextChanged
        dtPickTankTightnessTest_ValueChanged(sender, e)
    End Sub
    '    Private Sub dtPickTankTightnessTest_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickTankTightnessTest.EnabledChanged
    '        If bolLoading Then Exit Sub
    '        If Not dtPickTankTightnessTest.Enabled Then
    '            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickTankTightnessTest)
    '        End If
    '    End Sub
    Private Sub chkTankDrpTubeInvControl_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkTankDrpTubeInvControl.Click
        If bolLoading Then Exit Sub
        Try
            pTank.DropTube = chkTankDrpTubeInvControl.Checked
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickTankInstallerSigned_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickTankInstallerSigned.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickTankInstallerSigned)
            pTank.DateSigned = UIUtilsGen.GetDatePickerValue(dtPickTankInstallerSigned)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub lblLicenseeSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblLicenseeSearch.Click
        Try
            strFromCompanySearch = "TANK"
            oCompanySearch = New CompanySearch
            oCompanySearch.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        Finally
            oCompanySearch = Nothing
        End Try
    End Sub
    Private Sub dtPickLastUsed_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickLastUsed.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickLastUsed)
            pTank.DateLastUsed = UIUtilsGen.GetDatePickerValue(dtPickLastUsed)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmbTankClosureType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTankClosureType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.ClosureType = UIUtilsGen.GetComboBoxValueInt(cmbTankClosureType)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbTankInertFill_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTankInertFill.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pTank.InertMaterial = UIUtilsGen.GetComboBoxValueInt(cmbTankInertFill)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickSpillPreventionInstalled_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickSpillPreventionInstalled.ValueChanged
        btnTankSave.Enabled = True
    End Sub
    Private Sub dtPickSpillPreventionLastTested_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickSpillPreventionLastTested.ValueChanged
        btnTankSave.Enabled = True
    End Sub
    Private Sub dtPickOverfillPreventionLastInspected_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickOverfillPreventionLastInspected.ValueChanged
        btnTankSave.Enabled = True
    End Sub
    Private Sub dtPickSecondaryContainmentLastInspected_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickSecondaryContainmentLastInspected.ValueChanged
        btnTankSave.Enabled = True
    End Sub

    Private Sub dtPickOverfillPreventionInstalled_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickOverfillPreventionInstalled.ValueChanged
        Me.btnTankSave.Enabled = True
    End Sub

    Private Sub dtPickElectronicDeviceInspected_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickElectronicDeviceInspected.ValueChanged
        Me.btnTankSave.Enabled = True
    End Sub

    Private Sub dtPickATGLastInspected_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickATGLastInspected.ValueChanged
        Me.btnTankSave.Enabled = True
    End Sub

    Private Sub dtSheerValueTest_ValueChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtSheerValueTest.ValueChanged
        Me.btnPipeSave.Enabled = True
    End Sub

    Private Sub dtSecondaryContainmentInspected_ValueChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtSecondaryContainmentInspected.ValueChanged
        Me.btnPipeSave.Enabled = True
    End Sub
    Private Sub dtElectronicDeviceInspected_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtElectronicDeviceInspected.ValueChanged
        btnPipeSave.Enabled = True
    End Sub
    Private Sub lblTankDescDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankDescDisplay.Click
        ExpandCollapse(pnlTankDescriptionTop, lblTankDescDisplay)
    End Sub
    Private Sub lblTankDescHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankDescHead.Click
        ExpandCollapse(pnlTankDescriptionTop, lblTankDescDisplay)
    End Sub
    Private Sub lblTankInstallation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankInstallation.Click
        ExpandCollapse(pnlTankInstallation, lblTankInstallation)
    End Sub
    Private Sub lblDateofInstallation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblDateofInstallation.Click
        ExpandCollapse(pnlTankInstallation, lblTankInstallation)
    End Sub
    Private Sub lblTankTotalCapacity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankTotalCapacity.Click
        ExpandCollapse(pnlTankTotalCapacity, lblTankTotalCapacity)
    End Sub
    Private Sub lblTankTotalCapcityCaption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankTotalCapcityCaption.Click
        ExpandCollapse(pnlTankTotalCapacity, lblTankTotalCapacity)
    End Sub
    Private Sub lblTankMaterialDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankMaterialDisplay.Click
        ExpandCollapse(pnlTankMaterial, lblTankMaterialDisplay)
    End Sub
    Private Sub lblTankMaterialHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankMaterialHead.Click
        ExpandCollapse(pnlTankMaterial, lblTankMaterialDisplay)
    End Sub
    Private Sub lblTankReleaseDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankReleaseDisplay.Click
        ExpandCollapse(pnlTankRelease, lblTankReleaseDisplay)
    End Sub
    Private Sub lblTankReleaseHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankReleaseHead.Click
        ExpandCollapse(pnlTankRelease, lblTankReleaseDisplay)
    End Sub
    Private Sub lblTankInstallerOathDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankInstallerOathDisplay.Click
        ExpandCollapse(pnlTankInstallerOath, lblTankInstallerOathDisplay)
    End Sub
    Private Sub lblTankInstallerOath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankInstallerOath.Click
        ExpandCollapse(pnlTankInstallerOath, lblTankInstallerOathDisplay)
    End Sub
    Private Sub lblTankClosure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankClosure.Click
        ExpandCollapse(pnlTankClosure, lblTankClosure)
    End Sub
    Private Sub lblTankClosureCaption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTankClosureCaption.Click
        ExpandCollapse(pnlTankClosure, lblTankClosure)
    End Sub

    Private Sub btnAddPipe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPipe.Click
        Dim ugCompRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If pTank.IsDirty Or nTankID <= 0 Then
                MsgBox("Cannot Add Pipe before Saving the Tank")
                Exit Sub
            End If
            If chkTankCompartment.Checked Then
                ugCompRow = dGridCompartments.ActiveRow
                If ugCompRow Is Nothing Then
                    If dGridCompartments.Rows.Count = 1 Then
                        ugCompRow = dGridCompartments.Rows(0)
                    Else
                        MsgBox("Please select a valid row in the tank compartment grid prior to Adding Existing Pipes", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Invalid Selection")
                        Exit Sub
                    End If
                Else
                    If ugCompRow.Cells("COMPARTMENT NUMBER").Value <= 0 Then
                        MsgBox("Cannot Add Pipes to unsaved Compartments", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Invalid Selection")
                        Exit Sub
                    End If
                    If ugCompRow.Cells("COMPARTMENT NUMBER").Value <> nCompartmentNumber Then
                        nCompartmentNumber = ugCompRow.Cells("COMPARTMENT NUMBER").Value
                        pTank.Compartments.Retrieve(pTank.TankInfo, ugCompRow.Cells("ID").Value.ToString, False)
                    End If
                End If
            End If
            bolAddPipe = True
            ResetPipeID()
            ShowHideTankPipeScreen(False, True)
            If tbCntrlRegistration.SelectedTab.Name = tbPageManageTank.Name Then
                SetupTabs()
            Else
                tbCntrlRegistration.SelectedTab = tbPageManageTank
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnAddExistingPipe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddExistingPipe.Click
        Dim ugCompRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim frmAvailablePipes As AddAvailablePipes
        Try
            If pTank.IsDirty Or nTankID <= 0 Then
                MsgBox("Cannot Add Pipe before Saving the Tank")
                Exit Sub
            End If
            frmAvailablePipes = New AddAvailablePipes(pPipe, pTank)
            Me.Tag = "0"
            frmAvailablePipes.CallingForm = Me
            If chkTankCompartment.Checked Then
                ugCompRow = dGridCompartments.ActiveRow
                If ugCompRow Is Nothing Then
                    If dGridCompartments.Rows.Count = 1 Then
                        ugCompRow = dGridCompartments.Rows(0)
                    Else
                        MsgBox("Please select a valid row in the tank compartment grid prior to Adding Existing Pipes", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Invalid Selection")
                        Exit Sub
                    End If
                Else
                    If ugCompRow.Cells("COMPARTMENT NUMBER").Value <= 0 Then
                        MsgBox("Cannot Add Pipes to unsaved Compartments", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Invalid Selection")
                        Exit Sub
                    End If
                    If ugCompRow.Cells("COMPARTMENT NUMBER").Value <> nCompartmentNumber Then
                        nCompartmentNumber = ugCompRow.Cells("COMPARTMENT NUMBER").Value
                        pTank.Compartments.Retrieve(pTank.TankInfo, ugCompRow.Cells("ID").Value.ToString, False)
                    End If
                    frmAvailablePipes.Facility_id = nFacilityID
                    frmAvailablePipes.Tank_id = nTankID
                    frmAvailablePipes.Compartment_number = nCompartmentNumber
                End If
            Else
                frmAvailablePipes.Facility_id = nFacilityID
                frmAvailablePipes.Tank_id = nTankID
                frmAvailablePipes.Compartment_number = nCompartmentNumber
            End If
            frmAvailablePipes.ShowDialog()
            If Me.Tag = "1" Then
                PopulateTankPipeGrid(nFacilityID, False)
            End If
        Catch ex As Exception
            ShowError(ex)
        Finally
            If Not frmAvailablePipes Is Nothing Then frmAvailablePipes.Dispose()
        End Try
    End Sub
    Private Sub btnDetachPipes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDetachPipes.Click
        Dim frmDetachPipes As DetachPipes
        Try
            If pTank.IsDirty Or nTankID <= 0 Then
                MsgBox("Cannot Detach Pipe(s) before Saving the Tank")
                Exit Sub
            End If
            If IsNothing(frmDetachPipes) Then
                frmDetachPipes = New DetachPipes
            End If
            frmDetachPipes.Text = "Registration - Detach Pipes for Tank (" + pTank.TankIndex.ToString + ") - " + lblFacilityIDValue.Text + " (" + txtFacilityName.Text + ")"
            frmDetachPipes.FacilityID = nFacilityID
            frmDetachPipes.TankID = nTankID
            frmDetachPipes.TankObj = pTank
            frmDetachPipes.CallingForm = Me
            Me.Tag = "0"
            frmDetachPipes.ShowDialog()
            If Me.Tag = "1" Then
                ' refresh tank pipe grid
                PopulateTankPipeGrid(nFacilityID, False)
                PopulateTank(nTankID)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnTankSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTankSave.Click
        Dim bolAddTOSIActivity As Boolean = False
        Dim bolDeleteTOSIActivity As Boolean = False
        Dim bolUnregulatedTank As Boolean = False
        Dim facCAPStatus As Integer = pTank.FacilityInfo.CapStatus
        Try

            ' 2197
            If pOwn.Facilities.CAPCandidate Then
                If Not CheckCapDates(True) Then
                    Exit Sub
                End If
            End If
            ' #3032
            'If pTank.TankStatus = 429 And pTank.TankStatusOriginal <> 429 Then
            '    bolAddTOSIActivity = True
            'ElseIf pTank.TankStatusOriginal = 429 And pTank.TankStatus <> 429 Then
            '    bolDeleteTOSIActivity = True
            'End If

            If pTank.TankStatus = 430 Then
                bolUnregulatedTank = True
            End If

            If Me.dGridCompartments.Rows.Count > 1 And Me.chkTankCompartment.Checked And pTank.Compartments.Capacity = 0 And pTank.Compartments.Substance = 0 And pTank.Compartments.FuelTypeId = 0 And pTank.Compartments.CCERCLA = 0 Then
                dGridCompartments.Rows(dGridCompartments.Rows.Count - 1).Delete()
                pTank.Compartments.Remove(pTank.Compartments.ID)
            End If
            If chkTankCompartment.Checked Then
                If dGridCompartments.Rows.Count = 1 Then
                    If MsgBox("Only one compartment has been defined for this compartmentalized tank.  Are you certain you wish to save it?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Only One Compartment") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End If
            Dim success As Boolean = False
            If pTank.TankId <= 0 Then
                pTank.CreatedBy = MC.AppUser.ID
            Else
                pTank.ModifiedBy = MC.AppUser.ID
            End If
            returnVal = String.Empty
            success = pTank.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID, False, False, False, , IIf(chkBoxReplacementTank.Visible, chkBoxReplacementTank.Checked, False))
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            'Added by Hua Cao 09/05/2008
            'For tblReg_Tankplus. -	Date Spill Prevention Installed,	Date Spill Prevention Last Tested,
            'Date Overfill Prevention Last Inspected,	Date Secondary Containment Last Inspected.

            Dim LocalUserSettings As Microsoft.Win32.Registry
            Dim conSQLConnection As New SqlConnection
            Dim cmdSQLCommand As New SqlCommand
            Dim tankReader As SqlDataReader
            Dim tankReaderHasRows As Boolean

            conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")

            conSQLConnection.Open()
            cmdSQLCommand.Connection = conSQLConnection

            cmdSQLCommand.CommandText = "select * from tblReg_TankPlus where TankID = " + pTank.TankId.ToString
            tankReader = cmdSQLCommand.ExecuteReader()

            If Not tankReader.HasRows() Then
                cmdSQLCommand.CommandText = "insert into tblReg_TankPlus values(" + pTank.TankId.ToString + ",'" + dtPickSpillPreventionInstalled.Value.ToString + "', '" + dtPickSpillPreventionLastTested.Value.ToString + "', '" + dtPickOverfillPreventionLastInspected.Value.ToString + "', '" + dtPickSecondaryContainmentLastInspected.Value.ToString + "', '" + Me.dtPickElectronicDeviceInspected.Value.ToString + "', '" + Me.dtPickATGLastInspected.Value.ToString + "', '" + Me.dtPickOverfillPreventionInstalled.Value.ToString + "', 0)"
                tankReaderHasRows = False
            Else
                tankReaderHasRows = True
            End If
            tankReader.Close()
            If Not tankReaderHasRows Then
                cmdSQLCommand.ExecuteNonQuery()
            End If
            '    cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateSpillPreventionInstalled = '" + dtPickSpillPreventionInstalled.Value + "', DateSpillPreventionLastTested = '" + dtPickSpillPreventionLastTested.Value + "', DateOverfillPreventionLastInspected = '" + dtPickOverfillPreventionLastInspected.Value + "', DateSecondaryContainmentLastInspected = '" + dtPickSecondaryContainmentLastInspected.Value + "', DateElectronicDeviceInspected = '" + Me.dtPickElectronicDeviceInspected.Value + "', DateATGLastInspected = '" + Me.dtPickATGLastInspected.Value + "', DateOverfillPreventionInstalled = '" + Me.dtPickOverfillPreventionInstalled.Value + "' where TankID = " + nTankID.ToString
            'insert NULL if disabled date
            If dtPickSpillPreventionInstalled.Enabled And dtPickSpillPreventionInstalled.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateSpillPreventionInstalled = '" + dtPickSpillPreventionInstalled.Value + "' where TankID = " + nTankID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateSpillPreventionInstalled = NULL where TankID = " + nTankID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            If dtPickSpillPreventionLastTested.Enabled And Me.dtPickSpillPreventionLastTested.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateSpillPreventionLastTested = '" + dtPickSpillPreventionLastTested.Value + "' where TankID = " + nTankID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateSpillPreventionLastTested = NULL where TankID = " + nTankID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            If dtPickOverfillPreventionLastInspected.Enabled And dtPickOverfillPreventionLastInspected.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateOverfillPreventionLastInspected = '" + dtPickOverfillPreventionLastInspected.Value + "' where TankID = " + nTankID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateOverfillPreventionLastInspected = NULL where TankID = " + nTankID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            If dtPickSecondaryContainmentLastInspected.Enabled And dtPickSecondaryContainmentLastInspected.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateSecondaryContainmentLastInspected = '" + dtPickSecondaryContainmentLastInspected.Value + "' where TankID = " + nTankID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateSecondaryContainmentLastInspected = NULL where TankID = " + nTankID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            If dtPickElectronicDeviceInspected.Enabled And dtPickElectronicDeviceInspected.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateElectronicDeviceInspected = '" + dtPickElectronicDeviceInspected.Value + "' where TankID = " + nTankID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateElectronicDeviceInspected = NULL where TankID = " + nTankID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            If dtPickATGLastInspected.Enabled And dtPickATGLastInspected.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateATGLastInspected = '" + dtPickATGLastInspected.Value + "' where TankID = " + nTankID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateATGLastInspected = NULL where TankID = " + nTankID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            If dtPickOverfillPreventionInstalled.Enabled And dtPickOverfillPreventionInstalled.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateOverfillPreventionInstalled = '" + dtPickOverfillPreventionInstalled.Value + "' where TankID = " + nTankID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_TankPlus set DateOverfillPreventionInstalled = NULL where TankID = " + nTankID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()

            If success Then
                MsgBox("Tank Saved Successfully")
                If nTankID <= 0 Then
                    bolAddTank = False
                    ' registration activity
                    If pOwn.Facilities.SignatureOnNF Then
                        pOwn.Facilities.SignatureOnNF = False
                        pOwn.Facilities.ModifiedBy = MC.AppUser.ID
                        returnVal = String.Empty
                        pOwn.Facilities.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID, True, False)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                        returnVal = String.Empty
                        PutRegistrationActivity(nFacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.SignatureRequired)
                    End If
                    PutRegistrationActivity(pTank.TankId, UIUtilsGen.EntityTypes.Tank, UIUtilsGen.ActivityTypes.AddTank)
                End If
                nTankID = pTank.TankId
                If bolAddTOSIActivity Then
                    PutRegistrationActivity(nTankID, UIUtilsGen.EntityTypes.Tank, UIUtilsGen.ActivityTypes.TankStatusTOSI)
                ElseIf bolDeleteTOSIActivity Then
                    DeleteRegistrationActivity(nTankID, UIUtilsGen.EntityTypes.Tank, UIUtilsGen.ActivityTypes.TankStatusTOSI)
                End If

                If bolUnregulatedTank Then
                    Dim compNumber As Integer = 0
                    If Not ugTankRow Is Nothing Then
                        If Not ugTankRow.ChildBands(0) Is Nothing Then
                            If Not ugTankRow.ChildBands(0).Rows Is Nothing Then
                                If ugTankRow.ChildBands(0).Rows.Count > 0 Then
                                    For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugTankRow.ChildBands(0).Rows
                                        compNumber = IIf(ugRow.Cells("COMPARTMENT NUMBER").Value Is DBNull.Value, 0, ugRow.Cells("COMPARTMENT NUMBER").Value)
                                        nPipeID = ugRow.Cells("PIPE ID").Value
                                        pPipe.Retrieve(pTank.TankInfo, nTankID.ToString + "|" + compNumber.ToString + "|" + ugRow.Cells("PIPE ID").Value.ToString, , False)
                                        pPipe.PipeStatusDesc = 430
                                        pPipe.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID, False, False)
                                    Next
                                End If
                            End If
                        End If
                    End If
                End If

                ' refresh tank pipe grid
                PopulateTankPipeGrid(nFacilityID, False)
                PopulateTank(nTankID)
                ' if cap status was changed, load cap info on facility screen
                If facCAPStatus <> pTank.FacilityInfo.CapStatus Then
                    UIUtilsGen.PopulateFacilityCapInfo(Me, pOwn.Facilities)
                    UIUtilsGen.UpdateOwnerCAPInfo(Me, pOwn)
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnTankCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTankCancel.Click
        Try
            pTank.Reset()
            ' tank reset calls compartment reset
            'pTank.Compartments.Reset()
            PopulateTank(nTankID)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnCopyTankProfileToNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyTankProfileToNew.Click
        Try
            pTank.CopyTankProfile(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            ' registration activity
            If pOwn.Facilities.SignatureOnNF Then
                pOwn.Facilities.SignatureOnNF = False
                pOwn.Facilities.ModifiedBy = MC.AppUser.ID
                returnVal = String.Empty
                pOwn.Facilities.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID, True, False)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                PutRegistrationActivity(nFacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.SignatureRequired)
            End If
            PutRegistrationActivity(pTank.TankId, UIUtilsGen.EntityTypes.Tank, UIUtilsGen.ActivityTypes.AddTank)

            If pTank.TankStatus = UIUtilsGen.TankPipeStatus.TOSI Then
                PutRegistrationActivity(nTankID, UIUtilsGen.EntityTypes.Tank, UIUtilsGen.ActivityTypes.TankStatusTOSI)
            End If

            PopulateTankPipeGrid(nFacilityID, False)
            PopulateTank(pTank.TankId)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnDeleteTank_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteTank.Click
        Try
            If MsgBox("Are you sure you want to delete the Specified Tank?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Dim success As Boolean = False
                pTank.ModifiedBy = MC.AppUser.ID
                success = pTank.DeleteTank(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                If success Then
                    ' need to check for reg activity for existing tanks only
                    If nTankID > 0 Then
                        Dim oReg As New MUSTER.BusinessLogic.pRegistration
                        oReg.RetrieveByOwnerID(nOwnerID)
                        If oReg.ID > 0 Then
                            Dim deletedCount As Integer = 0
                            oReg.Activity.Col = oReg.Activity.RetrieveByRegID(oReg.ID)
                            For Each xRegActInfo As MUSTER.Info.RegistrationActivityInfo In oReg.Activity.Col.Values
                                If xRegActInfo.EntityId = nTankID And xRegActInfo.EntityType = UIUtilsGen.EntityTypes.Tank Then ' And xRegActInfo.ActivityDesc = UIUtilsGen.ActivityTypes.TankStatusTOSI
                                    xRegActInfo.Processed = True
                                    xRegActInfo.Deleted = True
                                    deletedCount += 1
                                End If
                            Next
                            oReg.Save()
                            If deletedCount = oReg.Activity.Col.Count Then
                                btnRegister.Visible = False
                            Else
                                btnRegister.Visible = True
                            End If
                        End If
                    End If
                    ResetTankID()
                    tbCntrlRegistration.SelectedTab = tbPageFacilityDetail
                    UIUtilsGen.PopulateFacilityCapInfo(Me, pOwn.Facilities)
                    UIUtilsGen.UpdateOwnerCAPInfo(Me, pOwn)
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnToPipe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToPipe.Click
        Try
            If nTankID = 0 Then
                MsgBox("Please select a Tank First")
                Exit Sub
            ElseIf nTankID < 0 Then
                MsgBox("Please save the Tank First")
                Exit Sub
            End If

            If Not ugTankRow Is Nothing Then
                If Not ugTankRow.ChildBands(0) Is Nothing Then
                    If Not ugTankRow.ChildBands(0).Rows Is Nothing Then
                        If ugTankRow.ChildBands(0).Rows.Count > 0 Then
                            ugPipeRow = ugTankRow.ChildBands(0).Rows(0)
                            ShowHideTankPipeScreen(False, True)
                            nCompartmentNumber = ugPipeRow.Cells("COMPARTMENT NUMBER").Value
                            nPipeID = ugPipeRow.Cells("PIPE ID").Value
                            If tbCntrlRegistration.SelectedTab.Name <> tbPageManageTank.Name Then
                                tbCntrlRegistration.SelectedTab = tbPageManageTank
                            Else
                                SetupTabs()
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region

#Region "Pipe"
    Friend Sub PopulatePipe(ByVal pipeID As Integer)
        Try
            ShowHideTankPipeScreen(False, True)
            nPipeID = pipeID

            SetPipeRow(nTankID, nCompartmentNumber, nPipeID)

            pTank = pOwn.Facilities.FacilityTanks
            If pTank.TankId <> nTankID Then
                pTank.RetrieveTank(nTankID)
            End If
            If pTank.Compartments.COMPARTMENTNumber <> nCompartmentNumber Or pTank.Compartments.TankId <> nTankID Then
                pTank.Compartments.Retrieve(pTank.TankInfo, nTankID.ToString + "|" + nCompartmentNumber.ToString, False)
            End If
            pPipe = pTank.Pipes
            pPipe.Retrieve(pTank.TankInfo, nTankID.ToString + "|" + nCompartmentNumber.ToString + "|" + nPipeID.ToString, , False)
            If pPipe.PipeID <= 0 Then
                nPipeID = pPipe.PipeID

                ClearPipeForm()

                FillPipeForm()

                Me.Text = "Registration - Manage Pipe (New) - " & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
                CommentsMaintenance(, , True, True)
            Else
                ClearPipeForm()

                FillPipeForm()

                Me.Text = "Registration - Manage Pipe (" + pPipe.Index.ToString + ") - " & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
                MC.FlagsChanged(nFacilityID, UIUtilsGen.EntityTypes.Facility, "Registration", "NONE", nPipeID, UIUtilsGen.EntityTypes.Pipe)
                CommentsMaintenance(, , True)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub SetPipeRow(ByVal tnkID As Integer, ByVal compNum As Integer, ByVal pipeID As Integer)
        Try
            SetTankRow(nTankID)
            If Not ugTankRow Is Nothing Then
                If ugPipeRow Is Nothing Then
                    If Not ugTankRow.ChildBands Is Nothing Then
                        For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugTankRow.ChildBands(0).Rows
                            If ugRow.Cells("COMPARTMENT NUMBER").Value = compNum And ugRow.Cells("PIPE ID").Value = pipeID Then
                                ugPipeRow = ugRow
                                Exit For
                            End If
                        Next
                    End If
                Else
                    If ugPipeRow.Cells("PIPE ID").Value <> pipeID Or ugPipeRow.Cells("COMPARTMENT NUMBER").Value <> compNum Then
                        ugPipeRow = Nothing
                        SetPipeRow(tnkID, compNum, pipeID)
                    End If
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub SetPipeSaveCancel(ByVal bolValue As Boolean)
        btnPipeSave.Enabled = bolValue
        btnPipeCancel.Enabled = bolValue
        If nPipeID > 0 Then
            btnDeletePipe.Enabled = True
            btnCopyPipeProfile.Enabled = True
        ElseIf bolValue Then
            btnDeletePipe.Enabled = True
        End If
    End Sub
    Private Sub ClearPipeForm()
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True

            lblPipeCompartmentIndex.Text = String.Empty
            lblPipeIndex.Text = String.Empty

            lblOwnerLastEditedBy.Text = "Last Edited By : "
            lblOwnerLastEditedOn.Text = "Last Edited On : "

            cmbPipeStatus.Enabled = False
            cmbPipeStatus.DataSource = Nothing

            dtPickPipeInstalled.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeInstalled)
            dtPickDatePipePlacedInService.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickDatePipePlacedInService)
            lblPipeSubstanceValue.Text = String.Empty
            lblPipeCerclaNoValue.Text = String.Empty
            lblPipeFuelTypeValue.Text = String.Empty

            cmbPipeMaterial.DataSource = Nothing
            cmbPipeMaterial.Enabled = False
            cmbPipeOptions.DataSource = Nothing
            cmbPipeOptions.Enabled = False
            cmbPipeCPType.DataSource = Nothing
            cmbPipeCPType.Enabled = False
            cmbPipeManufacturerID.DataSource = Nothing
            cmbPipeManufacturerID.Enabled = False
            chkEmergencyPowerPipe.Checked = False
            dtPickPipeCPInstalled.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeCPInstalled)
            dtPickPipeCPLastTest.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeCPLastTest)

            chkPipeSumpsAtDispenser.Enabled = False
            chkPipeSumpsAtDispenser.Checked = False
            chkPipeSumpsAtTank.Enabled = False
            chkPipeSumpsAtTank.Checked = False
            cmbPipeTerminationDispenserType.DataSource = Nothing
            cmbPipeTerminationDispenserType.Enabled = False
            cmbPipeTerminationDispenserCPType.DataSource = Nothing
            cmbPipeTerminationDispenserCPType.Enabled = False
            cmbPipeTerminationTankType.DataSource = Nothing
            cmbPipeTerminationTankType.Enabled = False
            cmbPipeTerminationTankCPType.DataSource = Nothing
            cmbPipeTerminationTankCPType.Enabled = False
            dtPickPipeTerminationCPInstalled.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTerminationCPInstalled)
            dtPickPipeTerminationCPLastTested.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTerminationCPLastTested)

            cmbPipeType.DataSource = Nothing
            cmbPipeType.Enabled = False

            cmbPipeReleaseDetection1.DataSource = Nothing
            cmbPipeReleaseDetection1.Enabled = False
            dtPickPipeTightnessTest.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTightnessTest)
            cmbPipeReleaseDetection2.DataSource = Nothing
            cmbPipeReleaseDetection2.Enabled = False
            dtPickPipeLeakDetectorTest.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLeakDetectorTest)

            dtPickPipeSigned.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeSigned)
            txtPipeLicensee.Text = String.Empty
            txtPipeLicensee.Enabled = False
            txtPipeCompanyName.Enabled = False
            txtPipeCompanyName.Text = String.Empty
            lblPipeLicenseeCompanySearch.Enabled = False

            dtPickPipeLastUsed.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLastUsed)
            lblPipeClosureRcvdDateValue.Text = String.Empty
            lblPipeClosedOnDate.Text = String.Empty
            cmbPipeClosureType.DataSource = Nothing
            cmbPipeClosureType.Enabled = False
            cmbPipeInertFill.DataSource = Nothing
            cmbPipeInertFill.Enabled = False

            btnPipeSave.Enabled = False
            btnPipeCancel.Enabled = False
            btnCopyPipeProfile.Enabled = False
            btnDeletePipe.Enabled = False
            btnToTank.Enabled = False
            btnPipeComments.Enabled = False
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub FillPipeForm()
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            If nPipeID <= 0 Then
                lblTankIDValue2.Text = pTank.TankIndex.ToString
                lblTankIDValue.Text = pTank.TankIndex.ToString

                lblPipeCompartmentIndex.Text = nCompartmentNumber.ToString
                lblPipeIndex.Text = pPipe.Index.ToString

                lblOwnerLastEditedBy.Text = "Last Edited By : "
                lblOwnerLastEditedOn.Text = "Last Edited On : "

                ' Status
                EnablePipeStatus(True)
                PopulatePipeStatus("ADD")
                pPipe.PipeStatusDesc = pTank.TankStatus
                UIUtilsGen.SetComboboxItemByValue(cmbPipeStatus, pPipe.PipeStatusDesc)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeStatus, pPipe.PipeStatusDesc, "PROPERTY_ID", "=") Then
                    pPipe.PipeStatusDesc = 0
                End If

                FillPipeFormFields(True)
            Else
                bolLoading = True

                lblTankIDValue2.Text = pTank.TankIndex.ToString
                lblTankIDValue.Text = pTank.TankIndex.ToString

                lblPipeCompartmentIndex.Text = nCompartmentNumber.ToString
                lblPipeIndex.Text = pPipe.Index.ToString

                lblOwnerLastEditedBy.Text = "Last Edited By : " & IIf(pPipe.ModifiedBy = String.Empty, pPipe.CreatedBy, pPipe.ModifiedBy)
                If Date.Compare(pPipe.ModifiedOn, dtNullDate) = 0 Then
                    lblOwnerLastEditedOn.Text = "Last Edited On : " & pPipe.CreatedOn.ToString
                Else
                    lblOwnerLastEditedOn.Text = "Last Edited On : " & pPipe.ModifiedOn.ToString
                End If

                ' Status
                EnablePipeStatus(True)
                PopulatePipeStatus("EDIT")
                UIUtilsGen.SetComboboxItemByValue(cmbPipeStatus, pPipe.PipeStatusDesc)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeStatus, pPipe.PipeStatusDesc, "PROPERTY_ID", "=") Then
                    pPipe.PipeStatusDesc = 0
                End If
                If pPipe.PipeStatusDesc <= 0 Then
                    cmbPipeStatus.SelectedIndex = -1
                    If cmbPipeStatus.SelectedIndex <> -1 Then
                        cmbPipeStatus.SelectedIndex = -1
                    End If
                End If
                ' -- If pipe type <> (Not equal) pressurized, disable Date ShearValue test 
                If Me.cmbPipeType.SelectedValue = 266 Then
                    Me.dtSheerValueTest.Enabled = True
                    Me.dtPickPipeLeakDetectorTest.Enabled = True
                    If Me.cmbPipeReleaseDetection1.SelectedValue = 243 Then
                        '	If Pipe Type =  (266)“Pressurized” and Release Detection Group 1 = (243)‘Continuous Interstitial Monitoring’ then 
                        ' -- All options are available for Group 2 
                        Dim strSQL As String
                        Dim dtReturn As DataTable
                        'Kevin Henderson's request, no long need "Continous" in group 2 10/28/08
                        If pPipe.PipeLD <> 0 Then
                            strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498' and PROPERTY_ID_PARENT = " + pPipe.PipeLD.ToString
                        Else
                            strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498'"
                        End If
                        dtReturn = pPipe.GetDataTable(strSQL)

                        Me.cmbPipeReleaseDetection2.DataSource = dtReturn

                        cmbPipeReleaseDetection2.DisplayMember = "PROPERTY_NAME"
                        cmbPipeReleaseDetection2.ValueMember = "PROPERTY_ID"

                    End If
                Else
                    UIUtilsGen.SetDatePickerValue(dtSheerValueTest, dtNullDate)
                    Me.dtSheerValueTest.Checked = False
                    UIUtilsGen.SetDatePickerValue(dtPickPipeLeakDetectorTest, dtNullDate)
                    Me.dtSheerValueTest.Enabled = False
                    Me.dtPickPipeLeakDetectorTest.Enabled = False
                End If
                'If Release Detection Group 1 <> (not equal) ‘Visual Interstitial Monitoring’ 242, disable Date Pipe SecondaryContainmentInspected;
                If Me.cmbPipeReleaseDetection1.SelectedValue = 242 Then
                    '    Me.dtSecondaryContainmentInspected.Enabled = True
                Else
                    '    Me.dtSecondaryContainmentInspected.Enabled = False
                End If
                'If Release Detection Group 1 <> (Not Equal) ‘Continuous Interstitial Monitoring’, disable Date Pipe ElectronicDeviceInspected
                If Me.cmbPipeReleaseDetection1.SelectedValue = 243 Then
                    '    Me.dtElectronicDeviceInspected.Enabled = True
                Else
                    '    Me.dtElectronicDeviceInspected.Enabled = False
                End If
                FillPipeFormFields(True)
            End If
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub FillPipeFormFields(ByVal populateCombo As Boolean)
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            If cmbPipeStatus.Text <> String.Empty Or pPipe.PipeStatusDesc > 0 Then
                dtPickPipeInstalled.Enabled = True
                UIUtilsGen.SetDatePickerValue(dtPickPipeInstalled, pPipe.PipeInstallDate)

                dtPickDatePipePlacedInService.Enabled = True
                UIUtilsGen.SetDatePickerValue(dtPickDatePipePlacedInService, pPipe.PlacedInServiceDate)

                lblPipeSubstanceValue.Text = pTank.Compartments.SubstanceDesc
                lblPipeCerclaNoValue.Text = pTank.Compartments.CERCLADesc
                lblPipeFuelTypeValue.Text = pTank.Compartments.FuelTypeIdDesc

                ' PipeMatDesc = Pipe Material
                ' PipeModDesc = Secondary Pipe Option
                cmbPipeMaterial.Enabled = True
                PopulatePipeMaterial()
                UIUtilsGen.SetComboboxItemByValue(cmbPipeMaterial, pPipe.PipeMatDesc)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeMaterial, pPipe.PipeMatDesc, "PROPERTY_ID", "=") Then
                    pPipe.PipeMatDesc = 0
                End If
                If pPipe.PipeMatDesc <= 0 Then
                    cmbPipeMaterial.SelectedIndex = -1
                    If cmbPipeMaterial.SelectedIndex <> -1 Then
                        cmbPipeMaterial.SelectedIndex = -1
                    End If
                End If

                cmbPipeOptions.Enabled = True
                PopulatePipeOptions(pPipe.PipeStatusDesc, pPipe.PipeMatDesc)
                UIUtilsGen.SetComboboxItemByValue(cmbPipeOptions, pPipe.PipeModDesc)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeOptions, pPipe.PipeModDesc, "PROPERTY_ID", "=") Then
                    pPipe.PipeModDesc = 0
                End If
                If pPipe.PipeModDesc <= 0 Then
                    cmbPipeOptions.SelectedIndex = -1
                    If cmbPipeOptions.SelectedIndex <> -1 Then
                        cmbPipeOptions.SelectedIndex = -1
                    End If
                End If

                ' Enable Field              Condition
                ' Pipe CP Type              Pipe Mod Desc (Pipe Sec Option) like 'Cathodically Protected'
                ' Pipe CP Installed         Pipe Mod Desc (Pipe Sec Option) like 'Cathodically Protected'
                ' Pipe CP Last Tested       Pipe Mod Desc (Pipe Sec Option) like 'Cathodically Protected'
                If cmbPipeOptions.Text.IndexOf("Cathodically Protected") > -1 Then
                    cmbPipeCPType.Enabled = True
                    PopulatePipeCPType()
                    UIUtilsGen.SetComboboxItemByValue(cmbPipeCPType, pPipe.PipeCPType)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeCPType, pPipe.PipeCPType, "PROPERTY_ID", "=") Then
                        pPipe.PipeCPType = 0
                    End If
                    If pPipe.PipeCPType <= 0 Then
                        cmbPipeCPType.SelectedIndex = -1
                        If cmbPipeCPType.SelectedIndex <> -1 Then
                            cmbPipeCPType.SelectedIndex = -1
                        End If
                    End If

                    dtPickPipeCPInstalled.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickPipeCPInstalled, pPipe.PipeCPInstalledDate)

                    dtPickPipeCPLastTest.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickPipeCPLastTest, pPipe.PipeCPTest)
                Else
                    cmbPipeCPType.Enabled = False
                    cmbPipeCPType.DataSource = Nothing

                    dtPickPipeCPInstalled.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeCPInstalled)
                    pPipe.PipeCPInstalledDate = dtNullDate

                    dtPickPipeCPLastTest.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeCPLastTest)
                    pPipe.PipeCPTest = dtNullDate
                End If

                cmbPipeManufacturerID.Enabled = True
                PopulatePipeManufacturer()
                UIUtilsGen.SetComboboxItemByValue(cmbPipeManufacturerID, pPipe.PipeManufacturer)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeManufacturerID, pPipe.PipeManufacturer, "PROPERTY_ID", "=") Then
                    pPipe.PipeManufacturer = 0
                End If
                If pPipe.PipeManufacturer <= 0 Then
                    cmbPipeManufacturerID.SelectedIndex = -1
                    If cmbPipeManufacturerID.SelectedIndex <> -1 Then
                        cmbPipeManufacturerID.SelectedIndex = -1
                    End If
                End If

                chkEmergencyPowerPipe.Checked = pTank.TankEmergen

                chkPipeSumpsAtDispenser.Enabled = True
                chkPipeSumpsAtDispenser.Checked = pPipe.ContainSumpDisp

                chkPipeSumpsAtTank.Enabled = True
                chkPipeSumpsAtTank.Checked = pPipe.ContainSumpTank

                cmbPipeTerminationDispenserType.Enabled = True
                PopulatePipeTermDispType()
                UIUtilsGen.SetComboboxItemByValue(cmbPipeTerminationDispenserType, pPipe.TermTypeDisp)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeTerminationDispenserType, pPipe.TermTypeDisp, "PROPERTY_ID", "=") Then
                    pPipe.TermTypeDisp = 0
                End If
                If pPipe.TermTypeDisp <= 0 Then
                    cmbPipeTerminationDispenserType.SelectedIndex = -1
                    If cmbPipeTerminationDispenserType.SelectedIndex <> -1 Then
                        cmbPipeTerminationDispenserType.SelectedIndex = -1
                    End If
                End If

                cmbPipeTerminationTankType.Enabled = True
                PopulatePipeTermTankType()
                UIUtilsGen.SetComboboxItemByValue(cmbPipeTerminationTankType, pPipe.TermTypeTank)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeTerminationTankType, pPipe.TermTypeTank, "PROPERTY_ID", "=") Then
                    pPipe.TermTypeTank = 0
                End If
                If pPipe.TermTypeTank <= 0 Then
                    cmbPipeTerminationTankType.SelectedIndex = -1
                    If cmbPipeTerminationTankType.SelectedIndex <> -1 Then
                        cmbPipeTerminationTankType.SelectedIndex = -1
                    End If
                End If

                ' Enable Field              Condition
                ' Pipe Term CP Installed    Pipe Term Type at Tank / Disp = 'Coated/Wrapped Cathodically Protected'
                ' Pipe Term CP Last Tested  Pipe Term Type at Tank / Disp = 'Coated/Wrapped Cathodically Protected'
                If cmbPipeTerminationDispenserType.Text.IndexOf("Coated/Wrapped Cathodically Protected") > -1 Then
                    cmbPipeTerminationDispenserCPType.Enabled = True
                    PopulatePipeTermDispCPType()
                    UIUtilsGen.SetComboboxItemByValue(cmbPipeTerminationDispenserCPType, pPipe.TermCPTypeDisp)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeTerminationDispenserCPType, pPipe.TermCPTypeDisp, "PROPERTY_ID", "=") Then
                        pPipe.TermCPTypeDisp = 0
                    End If
                    If pPipe.TermCPTypeDisp <= 0 Then
                        cmbPipeTerminationDispenserCPType.SelectedIndex = -1
                        If cmbPipeTerminationDispenserCPType.SelectedIndex <> -1 Then
                            cmbPipeTerminationDispenserCPType.SelectedIndex = -1
                        End If
                    End If
                    dtPickPipeTerminationCPInstalled.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickPipeTerminationCPInstalled, pPipe.TermCPInstalledDate)

                    dtPickPipeTerminationCPLastTested.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickPipeTerminationCPLastTested, pPipe.TermCPLastTested)
                Else
                    cmbPipeTerminationDispenserCPType.Enabled = False
                    cmbPipeTerminationDispenserCPType.DataSource = Nothing
                End If

                If cmbPipeTerminationTankType.Text.IndexOf("Coated/Wrapped Cathodically Protected") > -1 Then
                    cmbPipeTerminationTankCPType.Enabled = True
                    PopulatePipeTermTankCPType()
                    UIUtilsGen.SetComboboxItemByValue(cmbPipeTerminationTankCPType, pPipe.TermCPTypeTank)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeTerminationTankCPType, pPipe.TermCPTypeTank, "PROPERTY_ID", "=") Then
                        pPipe.TermCPTypeTank = 0
                    End If
                    If pPipe.TermCPTypeTank <= 0 Then
                        cmbPipeTerminationTankCPType.SelectedIndex = -1
                        If cmbPipeTerminationTankCPType.SelectedIndex <> -1 Then
                            cmbPipeTerminationTankCPType.SelectedIndex = -1
                        End If
                    End If
                    dtPickPipeTerminationCPInstalled.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickPipeTerminationCPInstalled, pPipe.TermCPInstalledDate)

                    dtPickPipeTerminationCPLastTested.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickPipeTerminationCPLastTested, pPipe.TermCPLastTested)
                Else
                    cmbPipeTerminationTankCPType.Enabled = False
                    cmbPipeTerminationTankCPType.DataSource = Nothing
                End If

                If Not (cmbPipeTerminationDispenserCPType.Enabled Or cmbPipeTerminationTankCPType.Enabled) Then
                    pPipe.TermCPInstalledDate = dtNullDate
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTerminationCPInstalled)
                    dtPickPipeTerminationCPInstalled.Enabled = False

                    pPipe.TermCPLastTested = dtNullDate
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTerminationCPLastTested)
                    dtPickPipeTerminationCPLastTested.Enabled = False
                End If

                cmbPipeType.Enabled = True
                If populateCombo = True Then
                    PopulatePipeType()
                End If
                UIUtilsGen.SetComboboxItemByValue(cmbPipeType, pPipe.PipeTypeDesc)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeType, pPipe.PipeTypeDesc, "PROPERTY_ID", "=") Then
                    pPipe.PipeTypeDesc = 0
                End If
                If pPipe.PipeTypeDesc <= 0 Then
                    cmbPipeType.SelectedIndex = -1
                    If cmbPipeType.SelectedIndex <> -1 Then
                        cmbPipeType.SelectedIndex = -1
                    End If
                End If
                If pPipe.PipeLD = 242 Then
                    Me.dtSecondaryContainmentInspected.Enabled = True
                Else
                    UIUtilsGen.SetDatePickerValue(Me.dtSecondaryContainmentInspected, dtNullDate)
                    Me.dtSecondaryContainmentInspected.Enabled = False
                End If

                If pPipe.PipeLD = 243 Then
                    Me.dtElectronicDeviceInspected.Enabled = True
                Else
                    UIUtilsGen.SetDatePickerValue(Me.dtElectronicDeviceInspected, dtNullDate)
                    Me.dtElectronicDeviceInspected.Enabled = False
                End If

                ' PipeLD = Release Detection 1
                ' Alld Type = Release Detection 2

                ' Enable Field              Condition
                ' Release Detection 1       Pipe Type = 'Pressurized'
                ' Release Detection 1       Pipe Type = 'U.S. Suction'
                ' LTT Date                  Release Detection 1 = 'Line Tightness Testing'

                ' Release Detection 2       Pipe Type = 'Pressurized' and (PipeLD <> 'Deferred')
                ' (ALLD Type)
                ' ALLD Test Date            Pipe Type = 'Pressurized' and (PipeLD <> 'Deferred')
                '                                        ALLD Type = 'Mechanical' and Pipe Status <> 'TOSI'
                If cmbPipeType.Text.IndexOf("Pressurized") > -1 Then
                    cmbPipeReleaseDetection1.Enabled = True
                    ' If populateCombo = True Then 'Removed by Hua Cao 03/02/2009
                    PopulatePipeReleaseDetection1(pPipe.PipeModDesc)
                    'End If
                    'removed by hua cao 09/26/2008
                    '    UIUtilsGen.SetComboboxItemByValue(cmbPipeReleaseDetection1, pPipe.PipeLD)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeReleaseDetection1, pPipe.PipeLD, "PROPERTY_ID", "=") Then
                        ' pPipe.PipeLD = 0
                    End If
                    If pPipe.PipeLD <= 0 Then
                        cmbPipeReleaseDetection1.SelectedIndex = -1
                        If cmbPipeReleaseDetection1.SelectedIndex <> -1 Then
                            cmbPipeReleaseDetection1.SelectedIndex = -1
                        End If
                    End If

                If cmbPipeReleaseDetection1.Text.IndexOf("Line Tightness Testing") > -1 Then
                    dtPickPipeTightnessTest.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickPipeTightnessTest, pPipe.LTTDate)
                ElseIf cmbPipeReleaseDetection1.Text.IndexOf("Electronic ALLD with 0.2") > -1 Then
                    dtPickPipeTightnessTest.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTightnessTest)
                    pPipe.LTTDate = dtNullDate
                    'Electronic
                    pPipe.ALLDType = 497
                ElseIf cmbPipeReleaseDetection1.Text.IndexOf("Continuous Interstitial Monitoring") > -1 Then
                    dtPickPipeTightnessTest.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTightnessTest)
                    pPipe.LTTDate = dtNullDate
                    'Continuous Interstitial Monitoring
                    'Removed hua cao 09/26/2008
                    '    pPipe.ALLDType = 498
                Else
                    dtPickPipeTightnessTest.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTightnessTest)
                    pPipe.LTTDate = dtNullDate
                End If

                'Retreive pipePlus date info  -  added by Hua Cao on 09/09/2008
                If pPipe.PipeID > 0 Then
                    Dim ds As DataTable = pTank.GetDataTable("tblReg_PipePlus where PipeID = " + pPipe.PipeID.ToString)
                    If Not (ds Is Nothing) Then
                        If Not ds.Rows(0).Item("DateSheerValueTest") Is System.DBNull.Value Then
                            dtSheerValueTest.Format = DateTimePickerFormat.Short
                            dtSheerValueTest.Value = ds.Rows(0).Item("DateSheerValueTest")
                            dtSheerValueTest.Checked = True
                        Else
                            UIUtilsGen.SetDatePickerValue(dtSheerValueTest, dtNullDate)
                        End If

                        If Not ds.Rows(0).Item("DateSecondaryContainmentInspect") Is System.DBNull.Value Then
                            dtSecondaryContainmentInspected.Format = DateTimePickerFormat.Short
                            dtSecondaryContainmentInspected.Value = ds.Rows(0).Item("DateSecondaryContainmentInspect")
                            dtSecondaryContainmentInspected.Checked = True
                        Else
                            UIUtilsGen.SetDatePickerValue(dtSecondaryContainmentInspected, dtNullDate)
                        End If
                        If Not ds.Rows(0).Item("DateElectronicDeviceInspect") Is System.DBNull.Value Then
                            dtElectronicDeviceInspected.Format = DateTimePickerFormat.Short
                            dtElectronicDeviceInspected.Value = ds.Rows(0).Item("DateElectronicDeviceInspect")
                            dtElectronicDeviceInspected.Checked = True
                        Else
                            UIUtilsGen.SetDatePickerValue(dtElectronicDeviceInspected, dtNullDate)
                        End If
                        UIUtilsGen.SetDatePickerValue(dtPickPipeLeakDetectorTest, pPipe.ALLDTestDate)
                        If pPipe.ALLDTestDate >= Convert.ToDateTime("1900-01-01") Then
                            dtPickPipeLeakDetectorTest.Format = DateTimePickerFormat.Short
                            Me.dtPickPipeLeakDetectorTest.Value = pPipe.ALLDTestDate
                            dtPickPipeLeakDetectorTest.Checked = True
                        Else
                            UIUtilsGen.SetDatePickerValue(dtPickPipeLeakDetectorTest, dtNullDate)
                        End If
                    End If

                End If
                If Not cmbPipeReleaseDetection1.Text.IndexOf("Deferred") > -1 Then
                    cmbPipeReleaseDetection2.Enabled = True
                    If populateCombo = True Then
                        PopulatePipeReleaseDetection2(, pPipe.PipeModDesc)
                    End If

                    UIUtilsGen.SetComboboxItemByValue(cmbPipeReleaseDetection2, pPipe.ALLDType)
                    If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeReleaseDetection2, pPipe.ALLDType, "PROPERTY_ID", "=") Then
                        pPipe.ALLDType = 0
                    End If
                    If pPipe.ALLDType <= 0 Then
                        cmbPipeReleaseDetection2.SelectedIndex = -1
                        If cmbPipeReleaseDetection2.SelectedIndex <> -1 Then
                            cmbPipeReleaseDetection2.SelectedIndex = -1
                        End If
                    End If

                    If cmbPipeReleaseDetection2.Text.IndexOf("Mechanical") > -1 Then
                        ' #2890 Disable ALLD Test Date if Pipe Status = TOSI
                        If pPipe.PipeStatusDesc = 429 Then
                            dtPickPipeLeakDetectorTest.Enabled = False
                            ' UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLeakDetectorTest)
                            '    pPipe.ALLDTestDate = dtNullDate
                        Else
                            dtPickPipeLeakDetectorTest.Enabled = True
                            UIUtilsGen.SetDatePickerValue(dtPickPipeLeakDetectorTest, pPipe.ALLDTestDate)
                        End If
                        UIUtilsGen.SetComboboxItemByValue(cmbPipeReleaseDetection1, pPipe.PipeLD)
                    ElseIf cmbPipeReleaseDetection2.Text.IndexOf("Electronic") > -1 Then
                        dtPickPipeLeakDetectorTest.Enabled = False
                        ' UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLeakDetectorTest)
                        '   pPipe.ALLDTestDate = dtNullDate
                        'Electronic ALLD with 0.2 Test
                        'Removed by Hua Cao 09/26/2008
                        ' pPipe.PipeLD = 246
                        UIUtilsGen.SetComboboxItemByValue(cmbPipeReleaseDetection1, pPipe.PipeLD)
                        If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeReleaseDetection1, pPipe.PipeLD, "PROPERTY_ID", "=") Then
                            '      pPipe.PipeLD = 0
                        End If
                        If pPipe.PipeLD <= 0 Then
                            cmbPipeReleaseDetection1.SelectedIndex = -1
                            If cmbPipeReleaseDetection1.SelectedIndex <> -1 Then
                                cmbPipeReleaseDetection1.SelectedIndex = -1
                            End If
                        End If
                    ElseIf cmbPipeReleaseDetection2.Text.IndexOf("Continuous") > -1 Then
                        dtPickPipeLeakDetectorTest.Enabled = False
                        'UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLeakDetectorTest)
                        '  pPipe.ALLDTestDate = dtNullDate
                        'Continuous Interstitial Monitoring
                        'removed hua cao 09/26/2008
                        'pPipe.PipeLD = 243
                        UIUtilsGen.SetComboboxItemByValue(cmbPipeReleaseDetection1, pPipe.PipeLD)
                        If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeReleaseDetection1, pPipe.PipeLD, "PROPERTY_ID", "=") Then
                            '   pPipe.PipeLD = 0
                        End If
                        If pPipe.PipeLD <= 0 Then
                            cmbPipeReleaseDetection1.SelectedIndex = -1
                            If cmbPipeReleaseDetection1.SelectedIndex <> -1 Then
                                cmbPipeReleaseDetection1.SelectedIndex = -1
                            End If
                        End If
                    Else
                        dtPickPipeLeakDetectorTest.Enabled = False
                        UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLeakDetectorTest)
                        pPipe.ALLDTestDate = dtNullDate
                    End If
                Else
                    cmbPipeReleaseDetection2.Enabled = False
                    cmbPipeReleaseDetection2.DataSource = Nothing
                    pPipe.ALLDType = 0

                    dtPickPipeLeakDetectorTest.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLeakDetectorTest)
                    pPipe.ALLDTestDate = dtNullDate
                End If
            ElseIf cmbPipeType.Text.IndexOf("U.S. Suction") > -1 Then
                cmbPipeReleaseDetection1.Enabled = True
                PopulatePipeReleaseDetection1(pPipe.PipeModDesc)
                UIUtilsGen.SetComboboxItemByValue(cmbPipeReleaseDetection1, pPipe.PipeLD)
                If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeReleaseDetection1, pPipe.PipeLD, "PROPERTY_ID", "=") Then
                    '  pPipe.PipeLD = 0
                End If
                If pPipe.PipeLD <= 0 Then
                    cmbPipeReleaseDetection1.SelectedIndex = -1
                    If cmbPipeReleaseDetection1.SelectedIndex <> -1 Then
                        cmbPipeReleaseDetection1.SelectedIndex = -1
                    End If
                End If

                If cmbPipeReleaseDetection1.Text.IndexOf("Line Tightness Testing") > -1 Then
                    dtPickPipeTightnessTest.Enabled = True
                    UIUtilsGen.SetDatePickerValue(dtPickPipeTightnessTest, pPipe.LTTDate)
                Else
                    dtPickPipeTightnessTest.Enabled = False
                    UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTightnessTest)
                    pPipe.LTTDate = dtNullDate
                End If

                cmbPipeReleaseDetection2.Enabled = False
                cmbPipeReleaseDetection2.DataSource = Nothing
                pPipe.ALLDType = 0

                dtPickPipeLeakDetectorTest.Enabled = False
                UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLeakDetectorTest)
                pPipe.ALLDTestDate = dtNullDate
            Else
                cmbPipeReleaseDetection1.Enabled = False
                '  cmbPipeReleaseDetection1.DataSource = Nothing
                '  pPipe.PipeLD = 0

                dtPickPipeTightnessTest.Enabled = False
                UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeTightnessTest)
                pPipe.LTTDate = dtNullDate

                cmbPipeReleaseDetection2.Enabled = False
                cmbPipeReleaseDetection2.DataSource = Nothing
                pPipe.ALLDType = 0

                dtPickPipeLeakDetectorTest.Enabled = False
                UIUtilsGen.CreateEmptyFormatDatePicker(dtPickPipeLeakDetectorTest)
                pPipe.ALLDTestDate = dtNullDate
            End If

            FillPipeInstallerOathLicensee()

            dtPickPipeLastUsed.Enabled = True
            UIUtilsGen.SetDatePickerValue(dtPickPipeLastUsed, pPipe.DateLastUsed)

            If Date.Compare(pPipe.DateClosureRecd, dtNullDate) = 0 Then
                lblPipeClosureRcvdDateValue.Text = String.Empty
            Else
                lblPipeClosureRcvdDateValue.Text = pPipe.DateClosureRecd.ToShortDateString
            End If
            If Date.Compare(pPipe.DateClosed, dtNullDate) = 0 Then
                lblPipeClosedOnDate.Text = String.Empty
            Else
                lblPipeClosedOnDate.Text = pPipe.DateClosed.ToShortDateString
            End If

            cmbPipeClosureType.Enabled = True
            cmbPipeInertFill.Enabled = True
            PopulatePipeClosureType()
            PopulatePipeInertFill()

            UIUtilsGen.SetComboboxItemByValue(cmbPipeClosureType, pPipe.ClosureType)
            If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeClosureType, pPipe.ClosureType, "PROPERTY_ID", "=") Then
                pPipe.ClosureType = 0
            End If
            If pPipe.ClosureType <= 0 Then
                cmbPipeClosureType.SelectedIndex = -1
                If cmbPipeClosureType.SelectedIndex <> -1 Then
                    cmbPipeClosureType.SelectedIndex = -1
                End If
            End If

            UIUtilsGen.SetComboboxItemByValue(cmbPipeInertFill, pPipe.InertMaterial)
            If Not UIUtilsGen.ComboBoxContainsValueSourceIsDataTable(cmbPipeInertFill, pPipe.InertMaterial, "PROPERTY_ID", "=") Then
                pPipe.InertMaterial = 0
            End If
            If pPipe.InertMaterial <= 0 Then
                cmbPipeInertFill.SelectedIndex = -1
                If cmbPipeInertFill.SelectedIndex <> -1 Then
                    cmbPipeInertFill.SelectedIndex = -1
                End If
            End If

            If cmbPipeStatus.Text.IndexOf("Currently In Use") > -1 Then
                dtPickPipeLastUsed.Enabled = False
                cmbPipeClosureType.Enabled = False
                cmbPipeInertFill.Enabled = False
            ElseIf cmbPipeStatus.Text.IndexOf("Permanently Out of Use") > -1 Then
                If pPipe.Pipe.POU And pPipe.Pipe.NonPre88 Then
                    dtPickLastUsed.Enabled = False
                    cmbPipeClosureType.Enabled = False
                    cmbPipeInertFill.Enabled = False
                End If
                If Date.Compare(pPipe.DateLastUsed, CDate("12/22/1988")) >= 0 Then
                    cmbPipeClosureType.Enabled = False
                    cmbPipeInertFill.Enabled = False
                End If
            End If
            End If

            If nPipeID > 0 Then
                btnCopyTankProfileToNew.Enabled = True
                btnPipeComments.Enabled = True
                btnToTank.Enabled = True
            Else
                btnCopyTankProfileToNew.Enabled = False
                btnPipeComments.Enabled = False
                btnToTank.Enabled = True
            End If

            If Me.cmbPipeType.SelectedValue = 266 Then
                Me.dtPickPipeLeakDetectorTest.Enabled = True
            Else
                Me.dtPickPipeLeakDetectorTest.Enabled = False
            End If
            If (Me.dtPickPipeInstalled.Value >= Convert.ToDateTime("10-01-2008")) And (populateCombo = True) Then
                'If pipe install date is greater or equal to 10/01/08, then
                '-- Release Detection Group 1 is filtered for like ‘%Interstitial Monitoring%’ 
                '-- Secondary Pipe Option is filtered for like ‘%Double Walled%’.
                Dim strSQL As String
                Dim dtReturn As DataTable
                If pPipe.PipeModDesc <> 0 Then
                    strSQL = "VPIPERELEASEDETECTIONTYPE where property_name like '%Interstitial Monitoring%' and PROPERTY_ID_PARENT = " + pPipe.PipeModDesc.ToString
                Else
                    strSQL = "VPIPERELEASEDETECTIONTYPE where property_name like '%Interstitial Monitoring%'"
                End If
                dtReturn = pPipe.GetDataTable(strSQL)
                Me.cmbPipeReleaseDetection1.DataSource = dtReturn
                cmbPipeReleaseDetection1.DisplayMember = "PROPERTY_NAME"
                cmbPipeReleaseDetection1.ValueMember = "PROPERTY_ID"
                Me.cmbPipeReleaseDetection1.SelectedValue = pPipe.PipeLD
                If pPipe.PipeMatDesc <> 0 Then
                    strSQL = "vPIPESECONDARYOPTIONTYPE where property_name like '%Double-Walled%' and PROPERTY_ID_PARENT = " + pPipe.PipeMatDesc.ToString
                Else
                    strSQL = "vPIPESECONDARYOPTIONTYPE where property_name like '%Double-Walled%'"
                End If
                dtReturn = pPipe.GetDataTable(strSQL)
                Me.cmbPipeOptions.DataSource = dtReturn
                cmbPipeOptions.DisplayMember = "PROPERTY_NAME"
                cmbPipeOptions.ValueMember = "PROPERTY_ID"
                cmbPipeOptions.SelectedValue = pPipe.PipeModDesc
            End If

            SetPipeSaveCancel(pPipe.IsDirty)

        Catch ex As Exception
            Throw ex
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub FillPipeInstallerOathLicensee()
        Try
            dtPickPipeSigned.Enabled = True
            UIUtilsGen.SetDatePickerValue(dtPickPipeSigned, pPipe.DateSigned)

            txtPipeLicensee.Enabled = True
            txtPipeCompanyName.Enabled = True
            lblPipeLicenseeCompanySearch.Enabled = True

            pLicensee.Retrieve(pPipe.LicenseeID)
            txtPipeLicensee.Text = pLicensee.Licensee_name

            pCompany.Retrieve(pPipe.ContractorID)
            txtPipeCompanyName.Text = pCompany.COMPANY_NAME
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub EnablePipeStatus(ByVal bolValue As Boolean)
        cmbPipeStatus.Enabled = bolValue
    End Sub
    Private Sub PopulatePipeStatus(ByVal Mode As String)
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True

            cmbPipeStatus.DataSource = pPipe.PopulatePipeStatus(Mode)
            cmbPipeStatus.DisplayMember = "PROPERTY_NAME"
            cmbPipeStatus.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tank Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub PopulatePipeMaterial()
        Try
            cmbPipeMaterial.DataSource = pPipe.PopulatePipeMaterialOfConstruction
            cmbPipeMaterial.DisplayMember = "PROPERTY_NAME"
            cmbPipeMaterial.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeOptions(Optional ByVal pipeStatus As Integer = 0, Optional ByVal pipeMatDesc As Integer = 0)
        Try
            cmbPipeOptions.DataSource = pPipe.PopulatePipeSecondaryOptionNew(pipeStatus, pipeMatDesc)
            cmbPipeOptions.DisplayMember = "PROPERTY_NAME"
            cmbPipeOptions.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeCPType()
        Try
            cmbPipeCPType.DataSource = pPipe.PopulatePipeCPType
            cmbPipeCPType.DisplayMember = "PROPERTY_NAME"
            cmbPipeCPType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeManufacturer()
        Try
            cmbPipeManufacturerID.DataSource = pPipe.PopulatePipeManufacturer
            cmbPipeManufacturerID.DisplayMember = "PROPERTY_NAME"
            cmbPipeManufacturerID.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeTermDispType()
        Try
            cmbPipeTerminationDispenserType.DataSource = pPipe.PopulatePipeTerminationDispenserType
            cmbPipeTerminationDispenserType.DisplayMember = "PROPERTY_NAME"
            cmbPipeTerminationDispenserType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeTermTankType()
        Try
            cmbPipeTerminationTankType.DataSource = pPipe.PopulatePipeTerminationTankType
            cmbPipeTerminationTankType.DisplayMember = "PROPERTY_NAME"
            cmbPipeTerminationTankType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeTermDispCPType()
        Try
            cmbPipeTerminationDispenserCPType.DataSource = pPipe.PopulatePipeTerminationDispenserCPType
            cmbPipeTerminationDispenserCPType.DisplayMember = "PROPERTY_NAME"
            cmbPipeTerminationDispenserCPType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeTermTankCPType()
        Try
            cmbPipeTerminationTankCPType.DataSource = pPipe.PopulatePipeTerminationTankCPType
            cmbPipeTerminationTankCPType.DisplayMember = "PROPERTY_NAME"
            cmbPipeTerminationTankCPType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeType()
        Try
            cmbPipeType.DataSource = pPipe.PopulatePipeType

            cmbPipeType.DisplayMember = "PROPERTY_NAME"
            cmbPipeType.ValueMember = "PROPERTY_ID"
            Me.cmbPipeType.SelectedValue = -1
            If Me.cmbPipeType.SelectedValue = 266 Then
                Me.dtSheerValueTest.Enabled = True
                Me.dtPickPipeLeakDetectorTest.Enabled = True
                If Me.cmbPipeReleaseDetection1.SelectedValue = 243 Then
                    '	If Pipe Type =  (266)“Pressurized” and Release Detection Group 1 = (243)‘Continuous Interstitial Monitoring’ then 
                    ' -- All options are available for Group 2 
                    Dim strSQL As String
                    Dim dtReturn As DataTable
                    'Kevin Henderson's request, no long need "Continous" in group 2 10/28/08
                    If pPipe.PipeLD <> 0 Then
                        strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498' and PROPERTY_ID_PARENT = " + pPipe.PipeLD.ToString
                    Else
                        strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498'"
                    End If
                    dtReturn = pPipe.GetDataTable(strSQL)

                    Me.cmbPipeReleaseDetection2.DataSource = dtReturn
                    cmbPipeReleaseDetection2.DisplayMember = "PROPERTY_NAME"
                    cmbPipeReleaseDetection2.ValueMember = "PROPERTY_ID"

                End If
            Else

                UIUtilsGen.SetDatePickerValue(Me.dtSheerValueTest, dtNullDate)
                Me.dtSheerValueTest.Enabled = False
                Me.dtPickPipeLeakDetectorTest.Enabled = False
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeReleaseDetection1(Optional ByVal pipeModDesc As Integer = 0)
        Try
            If Me.dtPickPipeInstalled.Value < Convert.ToDateTime("10-01-2008") Then
                cmbPipeReleaseDetection1.DataSource = pPipe.PopulatePipeReleaseDetection1(pipeModDesc)
                cmbPipeReleaseDetection1.DisplayMember = "PROPERTY_NAME"
                cmbPipeReleaseDetection1.ValueMember = "PROPERTY_ID"
                cmbPipeReleaseDetection1.SelectedValue = pPipe.PipeLD
            End If
            'If Release Detection Group 1 <> (not equal) ‘Visual Interstitial Monitoring’ 242, disable Date Pipe SecondaryContainmentInspected;
            If Me.cmbPipeReleaseDetection1.SelectedValue = 242 Then
                '         Me.dtSecondaryContainmentInspected.Enabled = True
            Else
                '        Me.dtSecondaryContainmentInspected.Enabled = False
            End If
            'If Release Detection Group 1 <> (Not Equal) ‘Continuous Interstitial Monitoring’, disable Date Pipe ElectronicDeviceInspected
            If Me.cmbPipeReleaseDetection1.SelectedValue = 243 Then
                '       Me.dtElectronicDeviceInspected.Enabled = True
            Else
                '      Me.dtElectronicDeviceInspected.Enabled = False
            End If

            If Me.cmbPipeType.SelectedValue = 266 Then
                Me.dtSheerValueTest.Enabled = True
                Me.dtPickPipeLeakDetectorTest.Enabled = True
                If Me.cmbPipeReleaseDetection1.SelectedValue = 243 Then
                    '	If Pipe Type =  (266)“Pressurized” and Release Detection Group 1 = (243)‘Continuous Interstitial Monitoring’ then 
                    ' -- All options are available for Group 2 
                    Dim strSQL As String
                    Dim dtReturn As DataTable
                    'Kevin Henderson's request, no long need "Continous" in group 2 10/28/08
                    If pPipe.PipeLD <> 0 Then
                        strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498' and PROPERTY_ID_PARENT = " + pPipe.PipeLD.ToString
                    Else
                        strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498'"
                    End If
                    dtReturn = pPipe.GetDataTable(strSQL)
                    Me.cmbPipeReleaseDetection2.DataSource = dtReturn
                    cmbPipeReleaseDetection2.DisplayMember = "PROPERTY_NAME"
                    cmbPipeReleaseDetection2.ValueMember = "PROPERTY_ID"

                End If
            Else

                UIUtilsGen.SetDatePickerValue(Me.dtSheerValueTest, dtNullDate)
                Me.dtSheerValueTest.Enabled = False
                Me.dtPickPipeLeakDetectorTest.Enabled = False
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeReleaseDetection2(Optional ByVal pipeLD As Integer = 0, Optional ByVal pipeModDesc As Integer = 0)
        Try
            cmbPipeReleaseDetection2.DataSource = pPipe.PopulatePipeReleaseDetection2(pipeLD, pipeModDesc)
            cmbPipeReleaseDetection2.DisplayMember = "PROPERTY_NAME"
            cmbPipeReleaseDetection2.ValueMember = "PROPERTY_ID"

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeClosureType()
        Try
            cmbPipeClosureType.DataSource = pPipe.PopulateClosureType
            cmbPipeClosureType.DisplayMember = "PROPERTY_NAME"
            cmbPipeClosureType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulatePipeInertFill()
        Try
            cmbPipeInertFill.DataSource = pPipe.PopulateTankPipeInertFill
            cmbPipeInertFill.DisplayMember = "PROPERTY_NAME"
            cmbPipeInertFill.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub CheckPipeStatus(ByVal oOldStatus As Integer, ByVal nNewStatus As Integer)
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            '424 - CIU
            '425 - TOS
            '426 - POU
            '428 - RegPending
            '429 - TOSI
            '430 - Unregulated
            bolLoading = True

            If nNewStatus = 429 Or nNewStatus = 425 Or nNewStatus = 426 Then ' TOSI, TOS, POU
                ' set focus to date last used
                dtPickPipeLastUsed.Focus()
            End If
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub

    Private Function GetPrevNextPipe(ByVal pipeID As Integer, ByVal getNext As Boolean) As Integer
        Try
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim retVal As Integer = pipeID
            Dim sl As New SortedList
            Dim slRow As New SortedList

            For Each ugRow In dgPipesAndTanks.Rows
                If Not ugRow.ChildBands Is Nothing Then
                    'If ugRow.Cells("TANK ID").Value = nTankID Then
                    For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugRow.ChildBands(0).Rows
                        sl.Add(ugChildRow.Cells("TANK ID").Text + "|" + _
                                ugChildRow.Cells("COMPARTMENT NUMBER").Text + "|" + _
                                ugChildRow.Cells("PIPE ID").Text, _
                                ugChildRow.Cells("TANK ID").Text + "|" + _
                                ugChildRow.Cells("COMPARTMENT NUMBER").Text + "|" + _
                                ugChildRow.Cells("PIPE ID").Text)
                        slRow.Add(ugChildRow.Cells("TANK ID").Text + "|" + _
                                ugChildRow.Cells("COMPARTMENT NUMBER").Text + "|" + _
                                ugChildRow.Cells("PIPE ID").Text, ugChildRow)
                    Next
                    'Exit For
                End If
            Next

            Dim strcompNumPipeID As String = GetPrevNext(sl, getNext, nTankID.ToString + "|" + nCompartmentNumber.ToString + "|" + pipeID.ToString)
            ugPipeRow = slRow.Item(strcompNumPipeID)
            ugTankRow = ugPipeRow.ParentRow
            nTankID = strcompNumPipeID.Split("|")(0)
            nCompartmentNumber = strcompNumPipeID.Split("|")(1)
            retVal = strcompNumPipeID.Split("|")(2)
            Return retVal
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Function

    Private Sub cmbPipeStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeStatus.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.PipeStatusDesc = UIUtilsGen.GetComboBoxValueInt(cmbPipeStatus)
            FillPipeFormFields(False)
            CheckPipeStatus(pPipe.PipeStatusDescOriginal, pPipe.PipeStatusDesc)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub dtPickPipeInstalled_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeInstalled.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeInstalled)
            pPipe.PipeInstallDate = UIUtilsGen.GetDatePickerValue(dtPickPipeInstalled)
            If Me.dtPickPipeInstalled.Value >= Convert.ToDateTime("10-01-2008") Then
                'If pipe install date is greater or equal to 10/01/08, then
                '-- Release Detection Group 1 is filtered for like ‘%Interstitial Monitoring%’ 
                '-- Secondary Pipe Option is filtered for like ‘%Double Walled%’.
                Dim strSQL As String
                Dim dtReturn As DataTable
                Dim dtreturn1 As DataTable
                If pPipe.PipeModDesc <> 0 Then
                    strSQL = "VPIPERELEASEDETECTIONTYPE where property_name like '%Interstitial Monitoring%' and PROPERTY_ID_PARENT = " + pPipe.PipeModDesc.ToString
                Else
                    strSQL = "VPIPERELEASEDETECTIONTYPE where property_name like '%Interstitial Monitoring%'"
                End If
                dtReturn = pPipe.GetDataTable(strSQL)
                Me.cmbPipeReleaseDetection1.DataSource = dtReturn
                cmbPipeReleaseDetection1.DisplayMember = "PROPERTY_NAME"
                cmbPipeReleaseDetection1.ValueMember = "PROPERTY_ID"

                If pPipe.PipeMatDesc <> 0 Then
                    strSQL = "vPIPESECONDARYOPTIONTYPE where property_name like '%Double-Walled%' and PROPERTY_ID_PARENT = " + pPipe.PipeMatDesc.ToString
                Else
                    strSQL = "vPIPESECONDARYOPTIONTYPE where property_name like '%Double-Walled%'"
                End If
                dtreturn1 = pPipe.GetDataTable(strSQL)
                Me.cmbPipeOptions.DataSource = dtreturn1
                cmbPipeOptions.DisplayMember = "PROPERTY_NAME"
                cmbPipeOptions.ValueMember = "PROPERTY_ID"


            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickDatePipePlacedInService_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickDatePipePlacedInService.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(dtPickDatePipePlacedInService)
            pPipe.PlacedInServiceDate = UIUtilsGen.GetDatePickerValue(dtPickDatePipePlacedInService)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmbPipeMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeMaterial.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.PipeMatDesc = UIUtilsGen.GetComboBoxValueInt(cmbPipeMaterial)
            If pPipe.PipeMatDescOriginal <> pPipe.PipeMatDesc Then
                pPipe.PipeManufacturer = 0
                pPipe.PipeCPType = 0
            End If
            FillPipeFormFields(False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbPipeOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeOptions.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.PipeModDesc = UIUtilsGen.GetComboBoxValueInt(cmbPipeOptions)
            Me.PopulatePipeReleaseDetection1(pPipe.PipeModDesc)
            Me.cmbPipeReleaseDetection1.SelectedValue = -1
            FillPipeFormFields(False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbPipeCPType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPipeCPType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.PipeCPType = UIUtilsGen.GetComboBoxValueInt(cmbPipeCPType)
            FillPipeFormFields(False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbPipeManufacturerID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPipeManufacturerID.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.PipeManufacturer = UIUtilsGen.GetComboBoxValueInt(cmbPipeManufacturerID)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickPipeCPInstalled_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeCPInstalled.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeCPInstalled)
            pPipe.PipeCPInstalledDate = UIUtilsGen.GetDatePickerValue(dtPickPipeCPInstalled)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dtPickPipeCPLastTest_ValueChanged()
    '    Dim dttemp, dtValidDate As Date
    '    Try
    '        dttemp = pPipe.PipeCPTest
    '        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
    '        dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
    '        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
    '        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
    '            dtPickPipeCPLastTest.Refresh()
    '            MsgBox("Pipe CP Last Tested must be greater than or equal to " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub dtPickPipeCPLastTest_ValueChangedHandler()
    '    Invoke(New HandlerDelegate(AddressOf dtPickPipeCPLastTest_ValueChanged))
    'End Sub
    Private Sub dtPickPipeCPLastTest_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeCPLastTest.CloseUp
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeCPLastTest)
            pPipe.PipeCPTest = UIUtilsGen.GetDatePickerValue(dtPickPipeCPLastTest)
            dtPickPipeCPLastTest.Tag = pPipe.PipeCPTest
            'Dim thread As New Threading.Thread(AddressOf dtPickPipeCPLastTest_ValueChangedHandler)
            'thread.Start()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickPipeCPLastTest_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeCPLastTest.TextChanged
        dtPickPipeCPLastTest_ValueChanged(sender, e)
    End Sub

    Private Sub chkPipeSumpsAtDispenser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPipeSumpsAtDispenser.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.ContainSumpDisp = chkPipeSumpsAtDispenser.Checked
            If chkPipeSumpsAtDispenser.Checked Then
                UIUtilsGen.SetComboboxItemByValue(cmbPipeTerminationDispenserType, 489)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkPipeSumpsAtTank_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPipeSumpsAtTank.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.ContainSumpTank = Me.chkPipeSumpsAtTank.Checked
            If chkPipeSumpsAtTank.Checked Then
                UIUtilsGen.SetComboboxItemByValue(cmbPipeTerminationTankType, 482)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbPipeTerminationDispenserType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeTerminationDispenserType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.TermTypeDisp = UIUtilsGen.GetComboBoxValueInt(cmbPipeTerminationDispenserType)
            FillPipeFormFields(False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbPipeTerminationDispenserCPType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPipeTerminationDispenserCPType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.TermCPTypeDisp = UIUtilsGen.GetComboBoxValueInt(cmbPipeTerminationDispenserCPType)
            FillPipeFormFields(False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbPipeTerminationTankType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeTerminationTankType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.TermTypeTank = UIUtilsGen.GetComboBoxValueInt(cmbPipeTerminationTankType)
            FillPipeFormFields(False)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbPipeTerminationTankCPType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPipeTerminationTankCPType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.TermCPTypeTank = UIUtilsGen.GetComboBoxValueInt(cmbPipeTerminationTankCPType)
            FillPipeFormFields(False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickPipeTerminationCPInstalled_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeTerminationCPInstalled.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeTerminationCPInstalled)
            pPipe.TermCPInstalledDate = UIUtilsGen.GetDatePickerValue(dtPickPipeTerminationCPInstalled)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dtPickPipeTerminationCPLastTested_ValueChanged()
    '    Dim dttemp, dtValidDate As Date
    '    Try
    '        dttemp = pPipe.TermCPLastTested
    '        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
    '        dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
    '        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
    '        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
    '            dtPickPipeTerminationCPLastTested.Refresh()
    '            MsgBox("Termination CP Last Tested must be greater than " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub dtPickPipeTerminationCPLastTested_ValueChangedHandler()
    '    Invoke(New HandlerDelegate(AddressOf dtPickPipeTerminationCPLastTested_ValueChanged))
    'End Sub
    Private Sub dtPickPipeTerminationCPLastTested_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeTerminationCPLastTested.CloseUp
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeTerminationCPLastTested)
            pPipe.TermCPLastTested = UIUtilsGen.GetDatePickerValue(dtPickPipeTerminationCPLastTested)
            dtPickPipeTerminationCPLastTested.Tag = pPipe.TermCPLastTested
            'Dim thread As New Threading.Thread(AddressOf dtPickPipeTerminationCPLastTested_ValueChangedHandler)
            'thread.Start()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickPipeTerminationCPLastTested_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeTerminationCPLastTested.TextChanged
        dtPickPipeTerminationCPLastTested_ValueChanged(sender, e)
    End Sub

    Private Sub cmbPipeType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.PipeTypeDesc = UIUtilsGen.GetComboBoxValueInt(cmbPipeType)
            ' -- If pipe type <> (Not equal) pressurized 266, disable Date shear value test;
            FillPipeFormFields(False)

            If Me.cmbPipeType.SelectedValue = 266 Then
                Me.dtSheerValueTest.Enabled = True
                Me.dtPickPipeLeakDetectorTest.Enabled = True
                'If Pipe Type = “Pressurized” and Release Detection Group 1 = ‘Continuous Interstitial Monitoring’ then 
                'All options are available for Group 2 
                ' If Me.cmbPipeReleaseDetection1.SelectedValue = 243 Then
                ' -- All options are available for Group 2 
                Dim strSQL As String
                Dim dtReturn As DataTable
                'Kevin Henderson's request, no long need "Continous" in group 2 10/28/08
                If pPipe.PipeLD <> 0 Then
                    strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498' and PROPERTY_ID_PARENT = " + pPipe.PipeLD.ToString
                Else
                    strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498'"
                End If
                dtReturn = pPipe.GetDataTable(strSQL, , True)
                Me.cmbPipeReleaseDetection2.DataSource = dtReturn
                cmbPipeReleaseDetection2.DisplayMember = "PROPERTY_NAME"
                cmbPipeReleaseDetection2.ValueMember = "PROPERTY_ID"
                cmbPipeReleaseDetection2.SelectedIndex = -1
                'End If
            Else
                UIUtilsGen.SetDatePickerValue(Me.dtSheerValueTest, dtNullDate)
                Me.dtSheerValueTest.Enabled = False
                Me.dtPickPipeLeakDetectorTest.Enabled = False
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmbPipeReleaseDetection1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeReleaseDetection1.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.PipeLD = UIUtilsGen.GetComboBoxValueInt(cmbPipeReleaseDetection1)
            If pPipe.ALLDType = 497 Or pPipe.ALLDType = 498 Then

                pPipe.ALLDType = 0
            End If

            'If Release Detection Group 1 <> (not equal) ‘Visual Interstitial Monitoring’ 242, disable Date Pipe SecondaryContainmentInspected;

            If Me.cmbPipeReleaseDetection1.SelectedValue = 242 Then
                Me.dtSecondaryContainmentInspected.Enabled = True
            Else
                UIUtilsGen.SetDatePickerValue(Me.dtSecondaryContainmentInspected, dtNullDate)
                Me.dtSecondaryContainmentInspected.Enabled = False
            End If
            'If Release Detection Group 1 <> (Not Equal) ‘Continuous Interstitial Monitoring’, disable Date Pipe ElectronicDeviceInspected
            If (Me.cmbPipeReleaseDetection1.SelectedValue = 243 Or Me.cmbPipeReleaseDetection1.SelectedValue = 242) Then
                Me.dtElectronicDeviceInspected.Enabled = True
                If Me.cmbPipeType.SelectedValue = 266 Then
                    'If Pipe Type = “Pressurized” and Release Detection Group 1 = ‘Continuous Interstitial Monitoring’ then 
                    ' -- All options are available for Group 2 
                    Dim strSQL As String
                    Dim dtReturn As DataTable
                    'Kevin Henderson's request, no long need "Continous" in group 2 10/28/08
                    If pPipe.PipeLD <> 0 Then
                        strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498' and PROPERTY_ID_PARENT = " + pPipe.PipeLD.ToString
                    Else
                        strSQL = "vPIPEAUTOMATICLINELEAKDECTIONTYPE where Property_ID <> '498'"
                    End If
                    dtReturn = pPipe.GetDataTable(strSQL)

                    Me.cmbPipeReleaseDetection2.DataSource = dtReturn

                    cmbPipeReleaseDetection2.DisplayMember = "PROPERTY_NAME"
                    cmbPipeReleaseDetection2.ValueMember = "PROPERTY_ID"
                End If
            Else
                UIUtilsGen.SetDatePickerValue(Me.dtElectronicDeviceInspected, dtNullDate)
                Me.dtElectronicDeviceInspected.Enabled = False
            End If
            Me.cmbPipeReleaseDetection2.SelectedValue = -1

        Catch ex As Exception
            '    ShowError(ex)
        End Try
        Try
            FillPipeFormFields(False)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dtPickPipeTightnessTest_ValueChanged()
    '    Dim dttemp, dtValidDate As Date
    '    Try
    '        dttemp = pPipe.LTTDate
    '        If pPipe.PipeTypeDesc = 268 Then
    '            dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
    '            dtValidDate = DateAdd(DateInterval.Year, -3, dtValidDate)
    '            dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
    '        ElseIf pPipe.PipeTypeDesc = 266 Then
    '            dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
    '            dtValidDate = DateAdd(DateInterval.Year, -1, dtValidDate)
    '            dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
    '        Else
    '            dtPickPipeTightnessTest.Refresh()
    '            Exit Sub
    '        End If
    '        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
    '            dtPickPipeTightnessTest.Refresh()
    '            MsgBox("Last Pipe Tightness Test must be greater than or equal to " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub dtPickPipeTightnessTest_ValueChangedHandler()
    '    Invoke(New HandlerDelegate(AddressOf dtPickPipeTightnessTest_ValueChanged))
    'End Sub
    Private Sub dtPickPipeTightnessTest_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeTightnessTest.CloseUp
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeTightnessTest)
            pPipe.LTTDate = UIUtilsGen.GetDatePickerValue(dtPickPipeTightnessTest)
            dtPickPipeTightnessTest.Tag = pPipe.LTTDate
            'Dim thread As New Threading.Thread(AddressOf dtPickPipeTightnessTest_ValueChangedHandler)
            'thread.Start()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickPipeTightnessTest_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeTightnessTest.TextChanged
        dtPickPipeTightnessTest_ValueChanged(sender, e)
    End Sub
    Private Sub cmbPipeReleaseDetection2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeReleaseDetection2.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            'modified by Hua Cao 10/29/08
            'If Me.cmbPipeReleaseDetection2.SelectedIndex <> 0 Then
            If Not (Me.cmbPipeReleaseDetection2.DataSource Is Nothing) Then
                cmbPipeReleaseDetection2.DisplayMember = "PROPERTY_NAME"
                cmbPipeReleaseDetection2.ValueMember = "PROPERTY_ID"
            End If
            pPipe.ALLDType = UIUtilsGen.GetComboBoxValueInt(cmbPipeReleaseDetection2)
            FillPipeFormFields(False)
            '  End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    'Private Sub dtPickPipeLeakDetectorTest_ValueChanged()
    '    Dim dttemp, dtValidDate As Date
    '    Try
    '        dttemp = pPipe.ALLDTestDate
    '        dtValidDate = Today.Month.ToString + "/1/" + Today.Year.ToString
    '        dtValidDate = DateAdd(DateInterval.Year, -1, dtValidDate)
    '        dtValidDate = DateAdd(DateInterval.Month, 3, dtValidDate)
    '        If Date.Compare(dttemp, dtTodayPlus90Days) > 0 Or Date.Compare(dtValidDate, dttemp) > 0 Then
    '            dtPickPipeLeakDetectorTest.Refresh()
    '            MsgBox("Automatic Line Leak Detector Test must be greater than or equal to " + dtValidDate.ToShortDateString + vbCrLf + " and less than or equal to " + dtTodayPlus90Days.ToShortDateString)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub dtPickPipeLeakDetectorTest_ValueChangedHandler()
    '    Invoke(New HandlerDelegate(AddressOf dtPickPipeLeakDetectorTest_ValueChanged))
    'End Sub
    Private Sub dtPickPipeLeakDetectorTest_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeLeakDetectorTest.CloseUp
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeLeakDetectorTest)
            pPipe.ALLDTestDate = UIUtilsGen.GetDatePickerValue(dtPickPipeLeakDetectorTest)
            dtPickPipeLeakDetectorTest.Tag = pPipe.ALLDTestDate
            'Dim thread As New Threading.Thread(AddressOf dtPickPipeLeakDetectorTest_ValueChangedHandler)
            'thread.Start()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub dtPickPipeLeakDetectorTest_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeLeakDetectorTest.TextChanged
        dtPickPipeLeakDetectorTest_ValueChanged(sender, e)
    End Sub

    Private Sub dtPickPipeSigned_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeSigned.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeSigned)
            pPipe.DateSigned = UIUtilsGen.GetDatePickerValue(dtPickPipeSigned)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub lblPipeLicenseeCompanySearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeLicenseeCompanySearch.Click
        Try
            strFromCompanySearch = "PIPE"
            oCompanySearch = New CompanySearch
            oCompanySearch.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        Finally
            oCompanySearch = Nothing
        End Try
    End Sub

    Private Sub dtPickPipeLastUsed_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickPipeLastUsed.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickPipeLastUsed)
            pPipe.DateLastUsed = UIUtilsGen.GetDatePickerValue(dtPickPipeLastUsed)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbPipeClosureType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPipeClosureType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.ClosureType = UIUtilsGen.GetComboBoxValueInt(cmbPipeClosureType)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbPipeInertFill_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPipeInertFill.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pPipe.InertMaterial = UIUtilsGen.GetComboBoxValueInt(cmbPipeInertFill)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub lblPipeDescDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeDescDisplay.Click
        ExpandCollapse(pnlPipeDescription, lblPipeDescDisplay)
    End Sub
    Private Sub lblPipeDescHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeDescHead.Click
        ExpandCollapse(pnlPipeDescription, lblPipeDescDisplay)
    End Sub
    Private Sub lblPipeDateOfInstallation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeDateOfInstallation.Click
        ExpandCollapse(pnlPipeDateOfInstallation, lblPipeDateOfInstallation)
    End Sub
    Private Sub lblPipeDateOfInstallationCaption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeDateOfInstallationCaption.Click
        ExpandCollapse(pnlPipeDateOfInstallation, lblPipeDateOfInstallation)
    End Sub
    Private Sub lblPipeMaterialDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeMaterialDisplay.Click
        ExpandCollapse(pnlPipeMaterial, lblPipeMaterialDisplay)
    End Sub
    Private Sub lblPipeMaterialHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeMaterialHead.Click
        ExpandCollapse(pnlPipeMaterial, lblPipeMaterialDisplay)
    End Sub
    Private Sub lblPipeType_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeType.Click
        ExpandCollapse(pnlPipeType, lblPipeType)
    End Sub
    Private Sub lblPipeTypeCaption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeTypeCaption.Click
        ExpandCollapse(pnlPipeType, lblPipeType)
    End Sub
    Private Sub lblPipeReleaseDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeReleaseDisplay.Click
        ExpandCollapse(pnlPipeRelease, lblPipeReleaseDisplay)
    End Sub
    Private Sub lblPipeReleaseHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeReleaseHead.Click
        ExpandCollapse(pnlPipeRelease, lblPipeReleaseDisplay)
    End Sub
    Private Sub lblPipeInstallerOathDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeInstallerOathDisplay.Click
        ExpandCollapse(pnlPipeInstallerOath, lblPipeInstallerOathDisplay)
    End Sub
    Private Sub lblPipeInstallerOath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeInstallerOath.Click
        ExpandCollapse(pnlPipeInstallerOath, lblPipeInstallerOathDisplay)
    End Sub
    Private Sub lblPipeClosure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeClosure.Click
        ExpandCollapse(pnlPipeClosure, lblPipeClosure)
    End Sub
    Private Sub lblPipeClosureCaption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblPipeClosureCaption.Click
        ExpandCollapse(pnlPipeClosure, lblPipeClosure)
    End Sub

    Private Sub btnPipeSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPipeSave.Click
        Dim facCAPStatus As Integer = pTank.FacilityInfo.CapStatus
        Try
            ' 2197
            If pOwn.Facilities.CAPCandidate Then
                If Not CheckCapDates(False) Then
                    Exit Sub
                End If
            End If
            Dim success As Boolean = False
            If pPipe.PipeID <= 0 Then
                pPipe.CreatedBy = MC.AppUser.ID
            Else
                pPipe.ModifiedBy = MC.AppUser.ID
            End If
            returnVal = String.Empty
            success = pPipe.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID, False, False)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            'added by Hua Cao 09/09/2008 for tblReg_pipePlus extra date fields


            Dim LocalUserSettings As Microsoft.Win32.Registry
            Dim conSQLConnection As New SqlConnection
            Dim cmdSQLCommand As New SqlCommand
            Dim pipeReader As SqlDataReader
            Dim pipeReaderHasRows As Boolean
            conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")

            conSQLConnection.Open()
            cmdSQLCommand.Connection = conSQLConnection
            cmdSQLCommand.CommandText = "select * from tblReg_PipePlus where PipeID = " + pPipe.PipeID.ToString
            pipeReader = cmdSQLCommand.ExecuteReader()

            If Not pipeReader.HasRows() Then
                cmdSQLCommand.CommandText = "insert into tblReg_PipePlus values(" + pPipe.PipeID.ToString + ", '" + dtSheerValueTest.Value.ToString + "', '" + dtSecondaryContainmentInspected.Value.ToString + "', '" + dtElectronicDeviceInspected.Value.ToString + "', 0)"
                pipeReaderHasRows = False
                'cmdSQLCommand.CommandText = "update tblReg_PipePlus set DateSheerValueTest = '" + dtSheerValueTest.Value + "', DateSecondaryContainmentInspect = '" + dtSecondaryContainmentInspected.Value + "', DateElectronicDeviceInspect = '" + dtElectronicDeviceInspected.Value + "' where PipeID = " + pPipe.PipeID.ToString
            Else
                pipeReaderHasRows = True
            End If
            pipeReader.Close()
            If Not pipeReaderHasRows Then
                cmdSQLCommand.ExecuteNonQuery()
            End If

            If Me.dtSheerValueTest.Enabled And dtSheerValueTest.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_PipePlus set DateSheerValueTest = '" + dtSheerValueTest.Value + "' where PipeID = " + pPipe.PipeID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_PipePlus set DateSheerValueTest = NULL where PipeID = " + pPipe.PipeID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            If Me.dtSecondaryContainmentInspected.Enabled And Me.dtSecondaryContainmentInspected.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_PipePlus set DateSecondaryContainmentInspect = '" + dtSecondaryContainmentInspected.Value + "' where PipeID = " + pPipe.PipeID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_PipePlus set DateSecondaryContainmentInspect = NULL where PipeID = " + pPipe.PipeID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()

            If Me.dtElectronicDeviceInspected.Enabled And Me.dtElectronicDeviceInspected.Format <> DateTimePickerFormat.Custom Then
                cmdSQLCommand.CommandText = "update tblReg_PipePlus set DateElectronicDeviceInspect = '" + dtElectronicDeviceInspected.Value + "' where PipeID = " + pPipe.PipeID.ToString
            Else
                cmdSQLCommand.CommandText = "update tblReg_PipePlus set DateElectronicDeviceInspect = NULL where PipeID = " + pPipe.PipeID.ToString
            End If
            cmdSQLCommand.ExecuteNonQuery()
            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()



            If success Then
                MsgBox("Pipe Saved Successfully")
                ' refresh tank pipe grid
                PopulateTankPipeGrid(nFacilityID, False)
                If nPipeID <= 0 Then
                    nPipeID = pPipe.PipeID
                    ' # 2824 If cal entry (upcoming pipe replacement) exists for the facility, delete the entry
                    Dim colCal As MUSTER.Info.CalendarCollection = MusterContainer.pCalendar.RetrieveByOtherID(0, 0, "Facility : " + nFacilityID.ToString + " Upcoming Pipe Replacement", "DESCRIPTION")
                    If Not colCal Is Nothing Then
                        For Each calInfo As MUSTER.Info.CalendarInfo In colCal.Values
                            MusterContainer.pCalendar.Add(calInfo)
                            MusterContainer.pCalendar.Deleted = True
                        Next
                        MusterContainer.pCalendar.Flush(UIUtilsGen.ModuleID.Registration, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    End If
                End If
                If facCAPStatus <> pPipe.FacCapStatus Then
                    pTank.FacCapStatus = pPipe.FacCapStatus
                    pOwn.Facilities.CapStatusOriginal = pPipe.FacCapStatus
                    pOwn.Facilities.CapStatus = pPipe.FacCapStatus
                    UIUtilsGen.PopulateFacilityCapInfo(Me, pOwn.Facilities)
                    UIUtilsGen.UpdateOwnerCAPInfo(Me, pOwn)
                End If
                PopulatePipe(nPipeID)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnPipeCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPipeCancel.Click
        Try
            pPipe.Reset()
            PopulatePipe(nPipeID)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnCopyPipeProfile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopyPipeProfile.Click
        Try
            pPipe.CopyPipeProfile(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            PopulateTankPipeGrid(nFacilityID, False)
            PopulatePipe(pPipe.PipeID)
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Copy Pipe Profile" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDeletePipe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeletePipe.Click
        Try
            If MsgBox("Deleting the pipe will disassociate all the tanks that the pipe was connected to" + vbCrLf + vbCrLf + "Are you sure you want to delete the Specified Pipe?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                pPipe.DeletePipe(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                ' check for facility cap status change
                Dim ds As DataSet = pOwn.RunSQLQuery("SELECT CAP_STATUS FROM tblREG_FACILITY WHERE FACILITY_ID = " + nFacilityID.ToString)
                If pOwn.Facilities.ID <> nFacilityID Then
                    pOwn.Facilities.Retrieve(pOwn.OwnerInfo, nFacilityID, , "FACILITY", , )
                End If
                Dim nCap As Integer = 0
                If Not ds.Tables(0).Rows(0)(0) Is DBNull.Value Then
                    nCap = ds.Tables(0).Rows(0)(0)
                End If
                If pOwn.Facilities.CapStatus <> nCap Then
                    pOwn.Facilities.CapStatusOriginal = nCap
                    pOwn.Facilities.CapStatus = nCap
                    UIUtilsGen.PopulateFacilityCapInfo(Me, pOwn.Facilities)
                    UIUtilsGen.UpdateOwnerCAPInfo(Me, pOwn)
                End If

                ResetTankID()
                ShowHideTankPipeScreen(False, False)
                If tbCntrlRegistration.SelectedTab.Name <> tbPageManageTank.Name Then
                    tbCntrlRegistration.SelectedTab = tbPageManageTank
                Else
                    SetupTabs()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot delete Pipe" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
        'P1 02/20/05 end
    End Sub
    Private Sub btnToTank_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnToTank.Click
        Try
            If Not ugTankRow Is Nothing Then
                ' if adding new pipe, need to save or cancel adding before navigation away from the screen
                If bolAddPipe Then
                    If tbCntrlRegistration.Tag = tbCntrlPipe.Name And nPipeID < 0 Then
                        If MsgBox("Do you want to save the new Pipe?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            pPipe.Remove(nTankID.ToString + "|" + nCompartmentNumber.ToString + "|" + nPipeID.ToString)
                            ResetPipeID()
                            bolAddPipe = False
                        Else
                            btnPipeSave.PerformClick()
                            ' if pipe not saved, do not navigate away from pipe screen
                            If nPipeID <= 0 Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                ShowHideTankPipeScreen(True, False)
                ResetCompartmentNumber()
                If tbCntrlRegistration.SelectedTab.Name <> tbPageManageTank.Name Then
                    tbCntrlRegistration.SelectedTab = tbPageManageTank
                Else
                    SetupTabs()
                End If
                ugPipeRow = Nothing
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region

#Region "Form Events"
    Private Sub Registration_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            If tbCntrlRegistration.SelectedTab.Name <> tbPageOwnerDetail.Name Then
                tbCntrlRegistration.SelectedTab = tbPageOwnerDetail
            End If
            tbCntrlRegistration.Tag = tbPageOwnerDetail.Name
            PopulateOwner(nOwnerID, False)

            If nFacilityID > 0 Then
                tbCntrlRegistration.SelectedTab = tbPageFacilityDetail
                tbCntrlRegistration.Tag = tbPageFacilityDetail.Name
                PopulateFacility(nFacilityID)
            End If
        Catch ex As Exception
            ShowError(ex)
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub Registration_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            If pOwn.colIsDirty() Then
                Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                If Results = MsgBoxResult.Yes Then
                    Dim success As Boolean = False
                    If pOwn.ID <= 0 Then
                        pOwn.CreatedBy = MC.AppUser.ID
                    Else
                        pOwn.ModifiedBy = MC.AppUser.ID
                    End If
                    returnVal = String.Empty
                    success = pOwn.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        UIUtilsGen.RemoveOwner(pOwn, Me)
                        Exit Sub
                    End If

                    If Not success Then
                        e.Cancel = True
                        'bolValidateSuccess = True
                        'bolDisplayErrmessage = True
                        Exit Sub
                    End If
                ElseIf Results = MsgBoxResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                End If
            End If
            UIUtilsGen.RemoveOwner(pOwn, Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Registration_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Registration")
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub Registration_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        'RaiseEvent EnableDisable(True)
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Registration")
        If lblOwnerIDValue.Text <> String.Empty Then ' And lblFacilityIDValue.Text = String.Empty Then
            pOwn.Retrieve(Me.lblOwnerIDValue.Text, "SELF")
        End If
    End Sub
#End Region

#Region "External Events"
    Public Sub evtOwnerErr(ByVal ErrMsg As String) Handles pOwn.evtOwnerErr
        If ErrMsg <> String.Empty Then
            MsgBox(ErrMsg)
        End If
    End Sub
    Public Sub OwnersChanged(ByVal bolstate As Boolean) Handles pOwn.evtOwnersChanged
        SetOwnerSaveCancel(bolstate)
    End Sub
    Public Sub OwnerChanged(ByVal bolstate As Boolean) Handles pOwn.evtOwnerChanged
        SetOwnerSaveCancel(bolstate)
    End Sub
    Public Sub ValidationErrors(ByVal FacID As Integer, ByVal MsgStr As String) Handles pOwn.evtValidationErr
        If MsgStr <> String.Empty Then
            MsgBox(MsgStr + " on FacilityID" + FacID.ToString)
        End If
    End Sub
    Public Sub PersonaChanged(ByVal bolState As Boolean) Handles pOwn.evtPersonaChanged
        SetPersonaSaveCancel(bolState)
    End Sub
    Public Sub PersonasChanged(ByVal bolstate As Boolean) Handles pOwn.evtPersonasChanged
        SetPersonaSaveCancel(bolstate)
    End Sub
    Public Sub OwnerFlagsChanged(ByVal entityID As Integer, ByVal entityType As Integer) Handles pOwn.FlagsChanged
        MC.FlagsChanged(entityID, entityType, "Registration", Me.Text)
    End Sub

    Public Sub FacilityChanged(ByVal bolstate As Boolean) Handles pOwn.evtFacilityChanged
        SetFacilitySaveCancel(bolstate)
    End Sub
    Public Sub FacilitiesChanged(ByVal bolstate As Boolean) Handles pOwn.evtFacilitiesChanged
        SetFacilitySaveCancel(bolstate)
    End Sub
    'Public Sub OwnFacilityCAPStatusChanged(ByVal BolValue As Boolean, ByVal facID As Integer) Handles pOwn.evtOwnFacilityCAPStatusChanged
    '    If facID = pOwn.Facilities.ID Then
    '        UIUtilsGen.PopulateFacilityCapInfo(Me, pOwn.Facilities)
    '    End If
    'End Sub

    Private Sub pTank_evtTankChanged(ByVal bolValue As Boolean) Handles pTank.evtTankChanged
        SetTankSaveCancel(bolValue)
    End Sub
    Private Sub pTank_evtTankErr(ByVal strMessage As String) Handles pTank.evtTankErr
        If strMessage <> String.Empty Then
            MsgBox(strMessage)
        End If
    End Sub
    Private Sub pTank_evtTanksChanged(ByVal bolValue As Boolean) Handles pTank.evtTanksChanged
        SetTankSaveCancel(bolValue)
    End Sub
    Private Sub pTank_evtTankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String) Handles pTank.evtTankValidationErr
        If strMessage <> String.Empty Then
            MsgBox(strMessage)
        End If
    End Sub

    Private Sub pPipe_evtPipeChanged(ByVal BolState As Boolean) Handles pPipe.evtPipeChanged
        SetPipeSaveCancel(BolState)
    End Sub
    Private Sub pPipe_evtPipeErr(ByVal StrMessage As String) Handles pPipe.evtPipeErr
        If StrMessage <> String.Empty Then
            MsgBox(StrMessage)
        End If
    End Sub
    Private Sub pPipe_evtPipesChanged(ByVal BolState As Boolean) Handles pPipe.evtPipesChanged
        SetPipeSaveCancel(BolState)
    End Sub
#End Region

#Region "Flags"
#Region "UI Support Routines"
    Private Sub SF_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles SF.FlagAdded
        MC.FlagsChanged(entityID, entityType, [Module], ParentFormText)
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        MC.RefreshCalendarInfo()
        MC.LoadDueToMeCalendar()
        MC.LoadToDoCalendar()
    End Sub
    Private Sub FlagMaintenance(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Select Case Me.tbCntrlRegistration.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    SF = New ShowFlags(nOwnerID, UIUtilsGen.EntityTypes.Owner, "Registration")
                Case tbPageFacilityDetail.Name
                    SF = New ShowFlags(nFacilityID, UIUtilsGen.EntityTypes.Facility, "Registration")
                Case Else
                    Exit Sub
            End Select
            SF.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub btnOwnerFlag_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerFlag.Click
        FlagMaintenance(sender, e)
    End Sub
    Private Sub btnFacFlags_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacFlags.Click
        FlagMaintenance(sender, e)
    End Sub
#End Region
#End Region

#Region "Comments"
    Private Sub btnOwnerComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerComment.Click
        CommentsMaintenance(sender, e)
    End Sub
    Private Sub btnFacComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacComments.Click
        CommentsMaintenance(sender, e)
    End Sub
    Private Sub btnTankComments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTankComments.Click
        CommentsMaintenance(sender, e)
    End Sub
    Private Sub btnPipeComments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPipeComments.Click
        CommentsMaintenance(sender, e)
    End Sub

    Private Sub CommentsMaintenance(Optional ByVal sender As System.Object = Nothing, Optional ByVal e As System.EventArgs = Nothing, Optional ByVal bolSetCounts As Boolean = False, Optional ByVal resetBtnColor As Boolean = False)
        Dim SC As ShowComments
        Dim nEntityType As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityName As String = String.Empty
        Dim oComments As MUSTER.BusinessLogic.pComments
        Dim bolEnableShowAllModules As Boolean = True
        Dim nCommentsCount As Integer = 0
        Try
            If Me.Text.StartsWith("Registration - Owner Summary") Or Me.Text.StartsWith("Registration - Owner Detail") Then
                nEntityType = UIUtilsGen.EntityTypes.Owner
                strEntityName = "Owner : " + nOwnerID.ToString + " " + Me.txtOwnerName.Text
                oComments = pOwn.Comments
                nEntityID = nOwnerID
            ElseIf Me.Text.StartsWith("Registration - Facility Detail") Then
                nEntityType = UIUtilsGen.EntityTypes.Facility
                strEntityName = "Facility : " + nFacilityID.ToString + " " + Me.txtFacilityName.Text
                oComments = pOwn.Facilities.Comments
                nEntityID = nFacilityID
            ElseIf Me.Text.StartsWith("Registration - Manage Tank") Then
                nEntityType = UIUtilsGen.EntityTypes.Tank
                strEntityName = "Tank Site ID : " + CStr(pTank.TankIndex)
                oComments = pOwn.Facilities.FacilityTanks.Comments
                nEntityID = nTankID
                bolEnableShowAllModules = False
            ElseIf Me.Text.StartsWith("Registration - Manage Pipe") Then
                nEntityType = UIUtilsGen.EntityTypes.Pipe
                strEntityName = "Pipe Site ID : " + CStr(pOwn.Facilities.FacilityTanks.Pipes.Index)
                oComments = pTank.Pipes.Comments
                nEntityID = nPipeID
                bolEnableShowAllModules = False
            ElseIf Me.Text.StartsWith("Registration - Owner/Facility Fees Report") Then
                Exit Sub
            Else : Exit Sub
            End If
            If Not resetBtnColor Then
                SC = New ShowComments(nEntityID, nEntityType, IIf(bolSetCounts, "", "Registration"), strEntityName, oComments, Me.Text, , bolEnableShowAllModules)
                If bolSetCounts Then
                    nCommentsCount = SC.GetCounts()
                Else
                    SC.ShowDialog()
                    nCommentsCount = SC.GetCounts()
                End If
            End If
            If nEntityType = UIUtilsGen.EntityTypes.Owner Then
                If nCommentsCount > 0 Then
                    btnOwnerComment.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnOwnerComment.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            ElseIf nEntityType = UIUtilsGen.EntityTypes.Facility Then
                If nCommentsCount > 0 Then
                    btnFacComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnFacComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            ElseIf nEntityType = UIUtilsGen.EntityTypes.Tank Then
                If nCommentsCount > 0 Then
                    btnTankComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnTankComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            ElseIf nEntityType = UIUtilsGen.EntityTypes.Pipe Then
                If nCommentsCount > 0 Then
                    btnPipeComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnPipeComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region

#Region "Contacts"
#Region "Owner Contacts"
    Private Sub ugOwnerContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugOwnerContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ugOwnerContacts)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub btnOwnerAddSearchContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAddSearchContact.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            objCntSearch = New ContactSearch(nOwnerID, UIUtilsGen.EntityTypes.Owner, "Registration", pConStruct)
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerModifyContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerModifyContact.Click
        Try
            ModifyContact(ugOwnerContacts)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerAssociateContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAssociateContact.Click
        Try
            AssociateContact(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerDeleteContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerDeleteContact.Click
        Try
            DeleteContact(ugOwnerContacts, nOwnerID)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOwnerShowActiveOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowActiveOnly.CheckedChanged
        'Dim dsContactsLocal As DataSet
        Try
            'If Not dsContacts Is Nothing Then
            SetOwnerFilter()
            'End If


        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOwnerShowContactsforAllModules_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowContactsforAllModules.CheckedChanged
        'Dim dsContactsLocal As DataSet
        Try


            'If Not dsContacts Is Nothing Then
            SetOwnerFilter()
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOwnerShowRelatedContacts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowRelatedContacts.CheckedChanged
        Try
            ' If Not dsContacts Is Nothing Then
            SetOwnerFilter()
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub SetOwnerFilter()
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strEntities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Try
            'strFilterString = String.Empty
            If chkOwnerShowActiveOnly.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If

            nEntityType = 9
            If chkOwnerShowContactsforAllModules.Checked Then
                ' User has the ability to view the contacts associated for the entity in other modules
                nEntityID = pOwn.ID
                nModuleID = 0

            Else
                nEntityID = pOwn.ID
                nModuleID = 612
            End If

            If chkOwnerShowRelatedContacts.Checked And strFacilityIdTags <> String.Empty Then
                strEntities = strFacilityIdTags
                nRelatedEntityType = 6
                'strFilterString += " OR ENTITYID IN (" + strFacilityIdTags + "))"
                'ElseIf chkOwnerShowRelatedContacts.Checked Then
                '    strFilterString += " OR (MODULEID = 612 And " + IIf(Not strFacilityIdTags = String.Empty, " ENTITYID in (" + strFacilityIdTags + ")))", "")
            Else
                strEntities = String.Empty
            End If

            dsContactsLocal = pConStruct.GetFilteredContacts(nEntityID, nModuleID, strEntities, bolActive, String.Empty, nEntityType, nRelatedEntityType)
            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            ugOwnerContacts.DataSource = dsContactsLocal





            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'strFilterString = String.Empty
            'If chkOwnerShowActiveOnly.Checked Then
            '    strFilterString = "(ACTIVE = 1"
            'Else
            '    strFilterString = "("
            'End If

            'If chkOwnerShowContactsforAllModules.Checked Then
            '    ' User has the ability to view the contacts associated for the entity in other modules
            '    If strFilterString = "(" Then
            '        strFilterString += "ENTITYID = " + pOwn.ID.ToString
            '    Else
            '        strFilterString += "AND ENTITYID = " + pOwn.ID.ToString
            '    End If
            'Else
            '    If strFilterString = "(" Then
            '        strFilterString += " MODULEID = 612 And ENTITYID = " + pOwn.ID.ToString
            '    Else
            '        strFilterString += " AND MODULEID = 612 And ENTITYID = " + pOwn.ID.ToString
            '    End If
            'End If

            'If chkOwnerShowRelatedContacts.Checked And strFacilityIdTags <> String.Empty Then
            '    strFilterString += " OR ENTITYID IN (" + strFacilityIdTags + "))"
            '    'ElseIf chkOwnerShowRelatedContacts.Checked Then
            '    '    strFilterString += " OR (MODULEID = 612 And " + IIf(Not strFacilityIdTags = String.Empty, " ENTITYID in (" + strFacilityIdTags + ")))", "")
            'Else
            '    strFilterString += ")"
            'End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugOwnerContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region
#Region "Facility Contacts"

    Private Sub ugFacilityContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFacilityContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ugFacilityContacts)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub btnFacilityAddSearchContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityAddSearchContact.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            objCntSearch = New ContactSearch(nFacilityID, UIUtilsGen.EntityTypes.Facility, "Registration", pConStruct)
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityModifyContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityModifyContact.Click
        Try
            ModifyContact(ugFacilityContacts)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityAssociateContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityAssociateContact.Click
        Try
            AssociateContact(ugFacilityContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityDeleteContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityDeleteContact.Click
        Try
            DeleteContact(ugFacilityContacts, nFacilityID)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkFacilityShowActiveContactOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFacilityShowActiveContactOnly.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFacilityFilter()
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkFacilityShowContactsforAllModule_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFacilityShowContactsforAllModule.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFacilityFilter()
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkFacilityShowRelatedContacts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFacilityShowRelatedContacts.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFacilityFilter()
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub SetFilter()
        Try
            strFilterString = String.Empty
            Dim strEntityID As String
            If tbCntrlRegistration.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                strEntityID = pOwn.ID.ToString
            Else
                strEntityID = pOwn.Facility.ID.ToString
            End If

            If chkOwnerShowActiveOnly.Checked Then
                strFilterString = "(ACTIVE = 1"
            Else
                strFilterString = "("
            End If

            If chkOwnerShowContactsforAllModules.Checked Then

                ' User has the ability to view the contacts associated for the entity in other modules
                If strFilterString = "(" Then
                    strFilterString += "ENTITYID = " + strEntityID
                Else
                    strFilterString += "AND ENTITYID = " + strEntityID
                End If
            Else
                If strFilterString = "(" Then
                    strFilterString += " MODULEID = 612 And ENTITYID = " + strEntityID
                Else
                    strFilterString += " AND MODULEID = 612 And ENTITYID = " + strEntityID
                End If
            End If
            If chkOwnerShowRelatedContacts.Checked Then
                strFilterString += " OR " + IIf(Not strFacilityIdTags = String.Empty, " ENTITYID in (" + strFacilityIdTags + "))", "")
            Else
                strFilterString += ")"
            End If

            dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            ugOwnerContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub SetFacilityFilter()
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strEntities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Dim strEntityAssocIDs As String = String.Empty
        Try

            'strFilterString = String.Empty
            If chkFacilityShowActiveContactOnly.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If
            nEntityType = 6

            If chkFacilityShowContactsforAllModule.Checked Then
                'User has the ability to view the contacts associated for the entity in other modules
                Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facilities.ID.ToString)
                strEntityAssocIDs = strFilterForAllModules
                nModuleID = 0
                nEntityID = pOwn.Facilities.ID.ToString

                nModuleID = 0
            Else
                nEntityID = pOwn.Facility.ID.ToString
                nModuleID = 612
            End If

            If chkFacilityShowRelatedContacts.Checked Then
                strEntities = strFacilityIdTags
                nRelatedEntityType = 6
            Else
                strEntities = String.Empty
            End If
            dsContactsLocal = pConStruct.GetFilteredContacts(nEntityID, nModuleID, strEntities, bolActive, strEntityAssocIDs, nEntityType, nRelatedEntityType)
            ugFacilityContacts.DataSource = dsContactsLocal

            'strFilterString = String.Empty
            'If chkFacilityShowActiveContactOnly.Checked Then
            '    strFilterString = "(ACTIVE = 1"
            'Else
            '    strFilterString = "("
            'End If


            'If chkFacilityShowContactsforAllModule.Checked Then
            '    ' User has the ability to view the contacts associated for the entity in other modules
            '    If strFilterString = "(" Then
            '        strFilterString += "ENTITYID = " + pOwn.Facility.ID.ToString
            '    Else
            '        strFilterString += "AND ENTITYID = " + pOwn.Facility.ID.ToString
            '    End If
            'Else
            '    If strFilterString = "(" Then
            '        strFilterString += " MODULEID = 612 And ENTITYID = " + pOwn.Facility.ID.ToString
            '    Else
            '        strFilterString += " AND MODULEID = 612 And ENTITYID = " + pOwn.Facility.ID.ToString
            '    End If
            'End If

            'If chkFacilityShowRelatedContacts.Checked Then
            '    strFilterString += " OR " + IIf(Not strFacilityIdTags = String.Empty, " ENTITYID in (" + strFacilityIdTags + "))", "")
            'Else
            '    strFilterString += ")"
            'End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugFacilityContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region
#Region "Common Functions"
    Private Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer)
        Dim dsContactsLocal As DataSet
        Try
            'dsContacts = pConStruct.GetAll()
            dsContactsLocal = pConStruct.GetFilteredContacts(EntityID, 612)
            'dsContacts.Tables(0).DefaultView.RowFilter = "MODULEID = 612 And ENTITYID = " + EntityID.ToString
            ugGrid.DataSource = dsContactsLocal.Tables(0).DefaultView 'dsContacts.Tables(0).DefaultView  

            '------ change column headings ---------------------------------------------------------
            ugGrid.DisplayLayout.Bands(0).Columns("CONTACT_Name").Header.Caption = "Contact Name"
            ugGrid.DisplayLayout.Bands(0).Columns("Address_One").Header.Caption = "Address 1"
            ugGrid.DisplayLayout.Bands(0).Columns("Address_Two").Header.Caption = "Address 2"
            ugGrid.DisplayLayout.Bands(0).Columns("AssocCompany").Header.Caption = "Assoc Company"
            ugGrid.DisplayLayout.Bands(0).Columns("Phone_Number_One").Header.Caption = "Phone 1"
            ugGrid.DisplayLayout.Bands(0).Columns("Ext_One").Header.Caption = "Ext 1"
            ugGrid.DisplayLayout.Bands(0).Columns("Phone_Number_Two").Header.Caption = "Phone 2"
            ugGrid.DisplayLayout.Bands(0).Columns("Ext_Two").Header.Caption = "Ext 2"
            ugGrid.DisplayLayout.Bands(0).Columns("CC_INFO").Header.Caption = "CC"
            ugGrid.DisplayLayout.Bands(0).Columns("IsPersonType").Header.Caption = "Person/Company"
            ugGrid.DisplayLayout.Bands(0).Columns("DATE_CREATED").Header.Caption = "Date Created"
            ugGrid.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Header.Caption = "Last Edited"
            ugGrid.DisplayLayout.Bands(0).Columns("DATEAssociated").Header.Caption = "Associated"
            ugGrid.DisplayLayout.Bands(0).Columns("Cell_Number").Header.Caption = "Cell Number"
            ugGrid.DisplayLayout.Bands(0).Columns("Fax_Number").Header.Caption = "Fax Number"
            ugGrid.DisplayLayout.Bands(0).Columns("Email_Address").Header.Caption = "Email Address"
            ugGrid.DisplayLayout.Bands(0).Columns("Email_Address_Personal").Header.Caption = "Personal Email"
            ugGrid.DisplayLayout.Bands(0).Columns("ACTIVE").Header.Caption = "Active"
            ugGrid.DisplayLayout.Bands(0).Columns("MODULE").Header.Caption = "Module"
            ugGrid.DisplayLayout.Bands(0).Columns("RELATIONSHIP").Header.Caption = "Relationship"

            '------ column widths ------------------------------------------------------------
            ugGrid.DisplayLayout.Bands(0).Columns("CONTACT_Name").Width = 120
            ugGrid.DisplayLayout.Bands(0).Columns("Address_One").Width = 100
            ugGrid.DisplayLayout.Bands(0).Columns("Address_Two").Width = 100
            ugGrid.DisplayLayout.Bands(0).Columns("City").Width = 90
            ugGrid.DisplayLayout.Bands(0).Columns("State").Width = 50
            ugGrid.DisplayLayout.Bands(0).Columns("Zip").Width = 70
            ugGrid.DisplayLayout.Bands(0).Columns("AssocCompany").Width = 110
            ugGrid.DisplayLayout.Bands(0).Columns("Phone_Number_One").Width = 100
            ugGrid.DisplayLayout.Bands(0).Columns("Ext_One").Width = 60
            ugGrid.DisplayLayout.Bands(0).Columns("Phone_Number_Two").Width = 100
            ugGrid.DisplayLayout.Bands(0).Columns("Ext_Two").Width = 60
            ugGrid.DisplayLayout.Bands(0).Columns("CC_INFO").Width = 40
            ugGrid.DisplayLayout.Bands(0).Columns("DATE_CREATED").Width = 80
            ugGrid.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Width = 80
            ugGrid.DisplayLayout.Bands(0).Columns("DATEAssociated").Width = 80

            '------ assign hidden columns ----------------------------------------------------
            ugGrid.DisplayLayout.Bands(0).Columns("Parent_Contact").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("Child_Contact").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ContactAssocID").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ModuleID").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ContactType").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ENTITYASSOCID").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("EntityType").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("LetterContactType").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("IsPerson").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("VENDOR_NUMBER").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("DISPLAYAS").Hidden = True
            '--------------------------------------------------------------------------------

            If tbCntrlRegistration.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                Me.chkOwnerShowActiveOnly.Checked = False
                Me.chkOwnerShowActiveOnly.Checked = True
            ElseIf tbCntrlRegistration.SelectedTab.Name.ToUpper = "TBPAGEFACILITYDETAIL" Then
                Me.chkFacilityShowActiveContactOnly.Checked = False
                Me.chkFacilityShowActiveContactOnly.Checked = True
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Function ModifyContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try
            If ugGrid.Rows.Count <= 0 Then Exit Function

            If ugGrid.ActiveRow Is Nothing Then
                MsgBox("Select row to Modify.")
                Exit Function
            End If
            Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow = ugGrid.ActiveRow
            If UCase(dr.Cells("Module").Value) <> "REGISTRATION" Then
                MsgBox(" Cannot modify " + dr.Cells("Module").Value.ToString + " Contacts in Registration")
                Exit Function
            End If
            If IsNothing(ContactFrm) Then
                ContactFrm = New Contacts(Integer.Parse(dr.Cells("EntityID").Value), CInt(dr.Cells("EntityType").Value), dr.Cells("Module").Value, CInt(dr.Cells("ContactID").Value), dr, pConStruct, "MODIFY")
            End If
            ContactFrm.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Function
    Private Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer)
        Try
            If ugGrid.Rows.Count <= 0 Then Exit Function

            If ugGrid.ActiveRow Is Nothing Then
                MsgBox("Select row to Associate.")
                Exit Function
            End If
            Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow = ugGrid.ActiveRow
            If ((CInt(ugGrid.ActiveRow.Cells("EntityID").Value) = nEntityID) And (CInt(ugGrid.ActiveRow.Cells("ModuleID").Value) = 612)) Then
                MsgBox("Selected contact is already associated with the current entity")
                Exit Function
            End If
            If IsNothing(ContactFrm) Then
                ContactFrm = New Contacts(nEntityID, nEntityType, "Registration", CInt(dr.Cells("ContactID").Value), dr, pConStruct, "ASSOCIATE")
            End If

            ContactFrm.ShowDialog()

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Function
    Private Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer)
        Try
            Dim result As DialogResult
            If ugGrid.Rows.Count <= 0 Then Exit Function

            If ugGrid.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete.")
                Exit Function
            End If
            If (CInt(ugGrid.ActiveRow.Cells("EntityID").Value) <> nEntityID) Or (CInt(ugGrid.ActiveRow.Cells("ModuleID").Value) <> 612) Then
                MsgBox("Selected contact is not associated with the current entity and cannot be deleted")
                Exit Function
            End If
            result = MessageBox.Show("Are you sure you wish to DELETE the record?", "MUSTER", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.No Then Exit Function

            pConStruct.Remove(ugGrid.ActiveRow.Cells("EntityAssocID").Text, CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal, MC.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Function
            End If
            ugGrid.ActiveRow.Delete(False)

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Function
    'Private Sub GetXHAndXLContacts()
    'Dim dtTable As DataTable
    'Dim i As Integer = 0
    'Try
    '    strXHContactName = String.Empty
    '    strXLContactName = String.Empty
    '    dtTable = pConStruct.GETContactName(nOwnerID, uiutilsgen.EntityTypes.Owner, UIUtilsGen.ModuleID.Registration)
    '    If dtTable.Rows.Count > 0 Then
    '        strXLContactName = dtTable.Rows(0).Item("CONTACT_Name")
    '    End If
    'Catch ex As Exception
    '    Throw ex
    'End Try
    'End Sub
    Private Enum ContactType
        XH = 1185
        XL = 1186
    End Enum
#End Region
#Region "Close Events"
    Private Sub Search_ContactAdded() Handles objCntSearch.ContactAdded
        If tbCntrlRegistration.SelectedTab.Name = tbPageOwnerDetail.Name Then
            LoadContacts(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
            SetOwnerFilter()
        Else
            LoadContacts(ugOwnerContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
            SetFacilityFilter()
        End If
    End Sub
    Private Sub Contact_ContactAdded() Handles ContactFrm.ContactAdded
        Try
            If tbCntrlRegistration.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL".ToUpper Then
                LoadContacts(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
            Else
                LoadContacts(ugOwnerContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub ContactFrm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContactFrm.Closing
        If Not ContactFrm Is Nothing Then
            ContactFrm = Nothing
        End If
    End Sub
    Private Sub objCntSearch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objCntSearch.Closing
        If Not objCntSearch Is Nothing Then
            objCntSearch = Nothing
        End If
    End Sub
#End Region
#End Region

    Private Sub btnRegister_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRegister.Click
        Dim oRLetter As New Reg_Letters
        Dim facColForSignatureRequired As New MUSTER.Info.FacilityCollection
        Dim facColFor2ndSignatureRequired As New MUSTER.Info.FacilityCollection
        Dim facColForUpcomingInstall As New MUSTER.Info.FacilityCollection
        'Dim facColForTOSIandAddTank As New MUSTER.Info.FacilityCollection
        Dim facInfo As MUSTER.Info.FacilityInfo
        Dim slRegActivity As New SortedList
        Dim alRegActivityAddTank As New ArrayList
        Dim alRegActivityTOSI As New ArrayList
        Dim alRegActivityFees As New ArrayList
        Dim alRegActivityTransferOwnership As New ArrayList
        Dim nInvalidRegActivityCount As Integer = 0

        Dim strFailureReason As String = String.Empty
        Dim strSignatureRequired As String = String.Empty
        Dim str2ndSignatureRequired As String = String.Empty
        Dim strUpcomingInstall As String = String.Empty
        Try

            'Check whether the registration process has rights or not.
            pOwn.CheckWriteAccessForRegistration(CType(UIUtilsGen.ModuleID.Registration, Integer), MC.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            ' if owner has no facilities, exit registration process
            If ugFacilityList.Rows Is Nothing Then
                strFailureReason = "The registration cannot be completed because : " + vbCrLf + "No facilities associated with owner: " + nOwnerID.ToString
            ElseIf Not ugFacilityList.Rows.Count > 0 Then
                strFailureReason = "The registration cannot be completed because : " + vbCrLf + "No facilities associated with owner: " + nOwnerID.ToString
            Else
                'GetXHAndXLContacts()

                CheckRegHeader(nOwnerID)

                If oReg.ID > 0 Then
                    oReg.Activity.Col = oReg.Activity.RetrieveByRegID(oReg.ID)
                    If oReg.Activity.Col.Count > 0 Then

                        For Each oregActivity As MUSTER.Info.RegistrationActivityInfo In oReg.Activity.Col.Values
                            If oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.AddOwner Then
                                If Not slRegActivity.Contains(UIUtilsGen.ActivityTypes.AddOwner) Then
                                    slRegActivity.Add(UIUtilsGen.ActivityTypes.AddOwner, oregActivity.EntityId.ToString)
                                End If
                            ElseIf oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.AddTank Then
                                If Not slRegActivity.Contains(UIUtilsGen.ActivityTypes.AddTank) Then
                                    slRegActivity.Add(UIUtilsGen.ActivityTypes.AddTank, oregActivity.EntityId.ToString)
                                Else
                                    slRegActivity.Item(UIUtilsGen.ActivityTypes.AddTank) += "," + oregActivity.EntityId.ToString
                                End If
                                If Not alRegActivityAddTank.Contains(oregActivity.EntityId) Then
                                    alRegActivityAddTank.Add(oregActivity.EntityId)
                                End If
                            ElseIf oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.SignatureRequired Then
                                If Not slRegActivity.Contains(UIUtilsGen.ActivityTypes.SignatureRequired) Then
                                    slRegActivity.Add(UIUtilsGen.ActivityTypes.SignatureRequired, oregActivity.EntityId.ToString)
                                Else
                                    slRegActivity.Item(UIUtilsGen.ActivityTypes.SignatureRequired) += "|" + oregActivity.EntityId.ToString
                                End If
                                facInfo = pOwn.Facilities.Retrieve(pOwn.OwnerInfo, oregActivity.EntityId, , "FACILITY", False, True)
                                If facInfo.ID > 0 Then
                                    strSignatureRequired += "Facility :" + facInfo.Name + " (" + facInfo.ID.ToString + ") requires signature on NF" + vbCrLf
                                    facColForSignatureRequired.Add(facInfo)
                                Else
                                    pOwn.OwnerInfo.facilityCollection.Remove(facInfo.ID)
                                    facInfo = New MUSTER.Info.FacilityInfo
                                End If
                            ElseIf oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.SecondLetterForSigRequired Then
                                If Not slRegActivity.Contains(UIUtilsGen.ActivityTypes.SecondLetterForSigRequired) Then
                                    slRegActivity.Add(UIUtilsGen.ActivityTypes.SecondLetterForSigRequired, oregActivity.EntityId.ToString)
                                Else
                                    slRegActivity.Item(UIUtilsGen.ActivityTypes.SecondLetterForSigRequired) += "|" + oregActivity.EntityId.ToString
                                End If
                                facInfo = pOwn.Facilities.Retrieve(pOwn.OwnerInfo, oregActivity.EntityId, , "FACILITY", False, True)
                                If facInfo.ID > 0 Then
                                    str2ndSignatureRequired += "Facility :" + facInfo.Name + " (" + facInfo.ID.ToString + ") requires 2nd signature on NF" + vbCrLf
                                    facColFor2ndSignatureRequired.Add(facInfo)
                                Else
                                    pOwn.OwnerInfo.facilityCollection.Remove(facInfo.ID)
                                    facInfo = New MUSTER.Info.FacilityInfo
                                End If
                            ElseIf oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.UpComingInstall Then
                                If Not slRegActivity.Contains(UIUtilsGen.ActivityTypes.UpComingInstall) Then
                                    slRegActivity.Add(UIUtilsGen.ActivityTypes.UpComingInstall, oregActivity.EntityId.ToString)
                                Else
                                    slRegActivity.Item(UIUtilsGen.ActivityTypes.UpComingInstall) += "|" + oregActivity.EntityId.ToString
                                End If
                                facInfo = pOwn.Facilities.Retrieve(pOwn.OwnerInfo, oregActivity.EntityId, , "FACILITY", False, True)
                                If facInfo.ID > 0 Then
                                    strUpcomingInstall += "Facility :" + facInfo.Name + " (" + facInfo.ID.ToString + ") requires Notice of Upcoming Installation" + vbCrLf
                                    facColForUpcomingInstall.Add(facInfo)
                                Else
                                    pOwn.OwnerInfo.facilityCollection.Remove(facInfo.ID)
                                    facInfo = New MUSTER.Info.FacilityInfo
                                End If
                            ElseIf oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.TankStatusTOSI Then
                                If Not slRegActivity.Contains(UIUtilsGen.ActivityTypes.TankStatusTOSI) Then
                                    slRegActivity.Add(UIUtilsGen.ActivityTypes.TankStatusTOSI, oregActivity.EntityId.ToString)
                                Else
                                    slRegActivity.Item(UIUtilsGen.ActivityTypes.TankStatusTOSI) += "," + oregActivity.EntityId.ToString
                                End If
                                If Not alRegActivityTOSI.Contains(oregActivity.EntityId) Then
                                    alRegActivityTOSI.Add(oregActivity.EntityId)
                                End If
                            ElseIf oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.Fees Then
                                If Not slRegActivity.Contains(UIUtilsGen.ActivityTypes.Fees) Then
                                    slRegActivity.Add(UIUtilsGen.ActivityTypes.Fees, oregActivity.EntityId.ToString)
                                Else
                                    slRegActivity.Item(UIUtilsGen.ActivityTypes.Fees) += "," + oregActivity.EntityId.ToString
                                End If
                                If Not alRegActivityFees.Contains(oregActivity.EntityId) Then
                                    alRegActivityFees.Add(oregActivity.EntityId)
                                End If
                            ElseIf oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.TransferOwnership Then
                                If Not slRegActivity.Contains(UIUtilsGen.ActivityTypes.TransferOwnership) Then
                                    slRegActivity.Add(UIUtilsGen.ActivityTypes.TransferOwnership, oregActivity.EntityId.ToString)
                                Else
                                    slRegActivity.Item(UIUtilsGen.ActivityTypes.TransferOwnership) += "," + oregActivity.EntityId.ToString
                                End If
                                If Not alRegActivityTransferOwnership.Contains(oregActivity.EntityId) Then
                                    alRegActivityTransferOwnership.Add(oregActivity.EntityId)
                                End If
                            Else
                                oregActivity.Processed = True
                                nInvalidRegActivityCount += 1
                            End If
                        Next

                    End If ' If oReg.Activity.Col.Count > 0 Then
                End If ' If oReg.ID > 0 Then
            End If

            If strFailureReason <> String.Empty Then
                MsgBox(strFailureReason)
                Exit Sub
            End If

            If strUpcomingInstall <> String.Empty Then
                If MsgBox(strUpcomingInstall + "Do you wish to produce Notice of Upcoming Installation letters for the above facilities?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Produce Upcoming Install Letter(s)?") = MsgBoxResult.Yes Then

                    oRLetter.GenerateUpcomingInstallLetter(nOwnerID, facColForUpcomingInstall, pOwn)
                    Dim strFacs As String = slRegActivity.Item(UIUtilsGen.ActivityTypes.UpComingInstall)
                    Dim al As New ArrayList
                    For Each fac As String In strFacs.Split("|")
                        If Not al.Contains(fac) Then al.Add(fac)
                    Next
                    For Each oregActivity As MUSTER.Info.RegistrationActivityInfo In oReg.Activity.Col.Values
                        If al.Contains(oregActivity.EntityId.ToString) And oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.UpComingInstall Then
                            oregActivity.Processed = True
                        End If
                    Next
                    oReg.Save()
                    slRegActivity.Remove(UIUtilsGen.ActivityTypes.UpComingInstall)
                End If
                Exit Sub
            End If

            If strSignatureRequired <> String.Empty Then
                If MsgBox("Registration cannot be completed because" + vbCrLf + strSignatureRequired + "Do you wish to produce the Letters for the above facilities?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Produce Signature Required Letter(s)?") = MsgBoxResult.No Then
                    Exit Sub
                Else
                    Try
                        oRLetter.GenerateSignatureRequiredLetter(nOwnerID, facColForSignatureRequired, pOwn)
                        Dim strFacs As String = slRegActivity.Item(UIUtilsGen.ActivityTypes.SignatureRequired)
                        Dim al As New ArrayList
                        For Each fac As String In strFacs.Split("|")
                            If Not al.Contains(fac) Then al.Add(fac)
                        Next
                        For Each oregActivity As MUSTER.Info.RegistrationActivityInfo In oReg.Activity.Col.Values
                            If al.Contains(oregActivity.EntityId.ToString) And oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.SignatureRequired Then
                                oregActivity.Processed = True
                            End If
                        Next
                        oReg.Save()
                        slRegActivity.Remove(UIUtilsGen.ActivityTypes.SignatureRequired)
                    Catch ex As Exception
                        ShowError(ex)
                        If Not ex.Message.StartsWith("Signature Required Letter for this owner ") Then
                            Exit Sub
                        End If
                    End Try
                End If
            End If

            If str2ndSignatureRequired <> String.Empty Then
                If MsgBox("Registration cannot be completed because" + vbCrLf + str2ndSignatureRequired + "Do you wish to produce the Letters for the above facilities?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Produce 2nd Signature Required Letter(s)?") = MsgBoxResult.No Then
                    Exit Sub
                Else
                    Try
                        oRLetter.GenerateSignatureRequiredLetter(nOwnerID, facColFor2ndSignatureRequired, pOwn, True)
                        Dim strFacs As String = slRegActivity.Item(UIUtilsGen.ActivityTypes.SecondLetterForSigRequired)
                        Dim al As New ArrayList
                        For Each fac As String In strFacs.Split("|")
                            If Not al.Contains(fac) Then al.Add(fac)
                        Next
                        For Each oregActivity As MUSTER.Info.RegistrationActivityInfo In oReg.Activity.Col.Values
                            If al.Contains(oregActivity.EntityId.ToString) And oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.SecondLetterForSigRequired Then
                                oregActivity.Processed = True
                            End If
                        Next
                        oReg.Save()
                        slRegActivity.Remove(UIUtilsGen.ActivityTypes.SecondLetterForSigRequired)
                    Catch ex As Exception
                        ShowError(ex)
                        If Not ex.Message.StartsWith("Signature Required Letter for this owner ") Then
                            Exit Sub
                        End If
                    End Try
                End If
            End If

            ' need to proceed, only if there are any other activities
            If Not slRegActivity Is Nothing Then
                If slRegActivity.Count > 0 Then
                    ' not doing transfer owner, - taken care after fac is transferred
                    ' not doing cap info needed

                    Dim strTransferOwnershipFacs As String = String.Empty
                    Dim strTanks As String = String.Empty
                    Dim strTOSITanks As String = String.Empty
                    Dim strFeeFacs As String = String.Empty
                    Dim alTanks As New ArrayList
                    Dim ds As DataSet
                    Dim tnk, fac As Integer
                    Dim bolDeleteNewOwnerActivity As Boolean = True

                    If alRegActivityTransferOwnership.Count > 0 Then
                        For Each fac In alRegActivityTransferOwnership.ToArray
                            strTransferOwnershipFacs += fac.ToString + ","
                        Next
                    End If
                    If alRegActivityAddTank.Count > 0 Then
                        For Each tnk In alRegActivityAddTank.ToArray
                            If Not alTanks.Contains(tnk) Then
                                alTanks.Add(tnk)
                                strTanks += tnk.ToString + ","
                            End If
                        Next
                    End If
                    If alRegActivityTOSI.Count > 0 Then
                        For Each tnk In alRegActivityTOSI.ToArray
                            If Not alTanks.Contains(tnk) Then
                                alTanks.Add(tnk)
                                strTanks += tnk.ToString + ","
                                strTOSITanks += tnk.ToString + ","
                            End If
                        Next
                    End If
                    If alRegActivityFees.Count > 0 Then
                        For Each fac In alRegActivityFees.ToArray
                            strFeeFacs += fac.ToString + ","
                        Next
                    End If

                    If strTransferOwnershipFacs <> String.Empty Then
                        strTransferOwnershipFacs = strTransferOwnershipFacs.Trim.TrimEnd(",")
                    End If
                    If strTanks <> String.Empty Then
                        strTanks = strTanks.Trim.TrimEnd(",")
                    End If
                    If strTOSITanks <> String.Empty Then
                        strTOSITanks = strTanks.Trim.TrimEnd(",")
                    End If
                    If strFeeFacs <> String.Empty Then
                        strFeeFacs = strFeeFacs.TrimEnd(",")
                    End If

                    'alTanks = New ArrayList

                    ' Abort process if Place In Service is Required for any tank which has activity (Add Tank / TOSI)
                    'If strTanks <> String.Empty Then
                    '    strTanks = strTanks.TrimEnd(",")
                    '    ds = pOwn.Facilities.CheckTankPlacedInService(strTanks)
                    '    If ds.Tables.Count > 0 Then
                    '        If ds.Tables(0).Rows.Count > 0 Then
                    '            For Each dr As DataRow In ds.Tables(0).Rows
                    '                If Not dr("TANK_INDEX") Is DBNull.Value Then
                    '                    If dr("TANK_INDEX").ToString <> String.Empty Then
                    '                        strFailureReason += dr("FACILITY_ID").ToString + vbTab + dr("TANK_INDEX").ToString
                    '                    End If
                    '                End If
                    '            Next
                    '            If strFailureReason <> String.Empty Then
                    '                strFailureReason = "The registration cannot be completed because : " + vbCrLf + _
                    '                    "The following facilities do not have Placed In Service Date for the given Tanks" + vbCrLf + _
                    '                    "Facility" + vbTab + "Tank(s)" + vbCrLf + strFailureReason
                    '                MsgBox(strFailureReason)
                    '                Exit Sub
                    '            End If
                    '        End If
                    '    End If
                    'End If

                    Dim alNewOwner As New ArrayList

                    If Not slRegActivity.Item(UIUtilsGen.ActivityTypes.AddOwner) Is Nothing Then
                        alNewOwner.Add(slRegActivity.Item(UIUtilsGen.ActivityTypes.AddOwner))
                    End If

                    ds = New DataSet

                    Dim strSQL As String = String.Empty

                    If strTanks <> String.Empty Then
                        strSQL = "(SELECT DISTINCT T.FACILITY_ID AS [ID], " + _
                            "F.[NAME], " + _
                            "(CASE WHEN LEN(RTRIM(LTRIM(A.ADDRESS_TWO))) > 0 THEN " + _
                            "(LTRIM(RTRIM(A.ADDRESS_LINE_ONE)) + ', ' + LTRIM(RTRIM(A.ADDRESS_TWO))) " + _
                            "ELSE (LTRIM(RTRIM(A.ADDRESS_LINE_ONE))) " + _
                            "END) AS ADDRESS, " + _
                            "LTRIM(RTRIM(A.CITY)) AS CITY, " + _
                            "LTRIM(RTRIM(A.STATE)) AS STATE, " + _
                            "LTRIM(RTRIM(A.ZIP)) AS ZIP " + _
                            "FROM TBLREG_TANK T LEFT OUTER JOIN TBLREG_FACILITY F ON F.FACILITY_ID = T.FACILITY_ID " + _
                            "LEFT OUTER JOIN tblREG_ADDRESS_MASTER A ON A.ADDRESS_ID = F.ADDRESS_ID " + _
                            "WHERE T.TANK_ID IN (" + strTanks + ")" + _
                            ")"
                    End If

                    If strFeeFacs <> String.Empty Then
                        If strSQL <> String.Empty Then
                            strSQL += " UNION "
                        End If
                        strSQL += "(SELECT DISTINCT F.FACILITY_ID AS [ID], " + _
                            "F.[NAME], " + _
                            "(CASE WHEN LEN(RTRIM(LTRIM(A.ADDRESS_TWO))) > 0 THEN " + _
                            "(LTRIM(RTRIM(A.ADDRESS_LINE_ONE)) + ', ' + LTRIM(RTRIM(A.ADDRESS_TWO))) " + _
                            "ELSE (LTRIM(RTRIM(A.ADDRESS_LINE_ONE))) " + _
                            "END) AS ADDRESS, " + _
                            "LTRIM(RTRIM(A.CITY)) AS CITY, " + _
                            "LTRIM(RTRIM(A.STATE)) AS STATE, " + _
                            "LTRIM(RTRIM(A.ZIP)) AS ZIP " + _
                            "FROM TBLREG_FACILITY F " + _
                            "LEFT OUTER JOIN tblREG_ADDRESS_MASTER A ON A.ADDRESS_ID = F.ADDRESS_ID " + _
                            "WHERE F.FACILITY_ID IN (" + strFeeFacs + ")" + _
                            ")"
                    End If

                    If strTransferOwnershipFacs <> String.Empty Then
                        If strSQL <> String.Empty Then
                            strSQL += " UNION "
                        End If
                        strSQL += "(SELECT DISTINCT F.FACILITY_ID AS [ID], " + _
                            "F.[NAME], " + _
                            "(CASE WHEN LEN(RTRIM(LTRIM(A.ADDRESS_TWO))) > 0 THEN " + _
                            "(LTRIM(RTRIM(A.ADDRESS_LINE_ONE)) + ', ' + LTRIM(RTRIM(A.ADDRESS_TWO))) " + _
                            "ELSE (LTRIM(RTRIM(A.ADDRESS_LINE_ONE))) " + _
                            "END) AS ADDRESS, " + _
                            "LTRIM(RTRIM(A.CITY)) AS CITY, " + _
                            "LTRIM(RTRIM(A.STATE)) AS STATE, " + _
                            "LTRIM(RTRIM(A.ZIP)) AS ZIP " + _
                            "FROM TBLREG_FACILITY F " + _
                            "LEFT OUTER JOIN tblREG_ADDRESS_MASTER A ON A.ADDRESS_ID = F.ADDRESS_ID " + _
                            "WHERE F.FACILITY_ID IN (" + strTransferOwnershipFacs + ")" + _
                            ")"
                    End If

                    If strSQL <> String.Empty Then
                        strSQL += " ORDER BY [ID]"
                        ds = pOwn.RunSQLQuery(strSQL)
                    End If

                    If ds.Tables.Count > 0 Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            ' create registration letter with all the required snippets
                            ' get tosi facilities from tosi tank id's
                            If strTOSITanks <> String.Empty Then
                                Dim dsTOSIFacs As DataSet = pOwn.RunSQLQuery("SELECT FACILITY_ID, TANK_INDEX FROM TBLREG_TANK WHERE TANK_ID IN (" + strTOSITanks + ") ORDER BY FACILITY_ID, TANK_INDEX")
                                Dim slTOSIFacs As New SortedList
                                Dim dr As DataRow
                                strTOSITanks = String.Empty
                                If dsTOSIFacs.Tables(0).Rows.Count > 0 Then
                                    For i As Integer = 0 To dsTOSIFacs.Tables(0).Rows.Count - 1
                                        dr = dsTOSIFacs.Tables(0).Rows(i)
                                        If Not slTOSIFacs.Contains(dr("FACILITY_ID")) Then
                                            slTOSIFacs.Add(dr("FACILITY_ID"), dr("TANK_INDEX").ToString)
                                        Else
                                            slTOSIFacs.Item(dr("FACILITY_ID")) += ", T" + dr("TANK_INDEX").ToString
                                        End If
                                    Next
                                    strTOSITanks = String.Empty
                                    For i As Integer = 0 To slTOSIFacs.Count - 1
                                        strTOSITanks += "Facility# " + slTOSIFacs.GetKey(i).ToString + " (" + slTOSIFacs.Item(slTOSIFacs.GetKey(i)).ToString + "), "
                                    Next
                                    strTOSITanks = strTOSITanks.Trim.TrimEnd(",")
                                End If
                            End If
                            oRLetter.GenerateRegistrationLetter(nOwnerID, slRegActivity, ds.Tables(0), pOwn, strTOSITanks, strTransferOwnershipFacs, alRegActivityTransferOwnership.Count)
                            If pOwn.OwnerL2CSnippet Then
                                pOwn.OwnerL2CSnippet = False
                                returnVal = String.Empty
                                pOwn.Save(UIUtilsGen.ModuleID.Registration, MC.AppUser.UserKey, returnVal, MC.AppUser.ID, True, False)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If
                            End If
                        Else
                            MsgBox("No Facilities have new Tanks or Tank Status changed to TOSI or Fees or Transfer Ownership")
                            bolDeleteNewOwnerActivity = False
                        End If
                    Else
                        MsgBox("No Facilities have new Tanks or Tank Status changed to TOSI or Fees or Transfer Ownership")
                        bolDeleteNewOwnerActivity = False
                    End If

                    For Each oregActivity As MUSTER.Info.RegistrationActivityInfo In oReg.Activity.Col.Values
                        If (alRegActivityAddTank.Contains(oregActivity.EntityId) And oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.AddTank) Or _
                            (alRegActivityTOSI.Contains(oregActivity.EntityId) And oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.TankStatusTOSI) Or _
                            (alRegActivityFees.Contains(oregActivity.EntityId) And oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.Fees) Or _
                            (alRegActivityTransferOwnership.Contains(oregActivity.EntityId) And oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.TransferOwnership) Or _
                            (bolDeleteNewOwnerActivity And (alNewOwner.Contains(oregActivity.EntityId.ToString) And oregActivity.ActivityDesc = UIUtilsGen.ActivityTypes.AddOwner)) Then
                            oregActivity.Processed = True
                        End If
                    Next

                    If slRegActivity.Contains(UIUtilsGen.ActivityTypes.TankStatusTOSI) Then
                        slRegActivity.Remove(UIUtilsGen.ActivityTypes.TankStatusTOSI)
                    End If
                    If slRegActivity.Contains(UIUtilsGen.ActivityTypes.AddTank) Then
                        slRegActivity.Remove(UIUtilsGen.ActivityTypes.AddTank)
                    End If
                    If slRegActivity.Contains(UIUtilsGen.ActivityTypes.AddOwner) Then
                        slRegActivity.Remove(UIUtilsGen.ActivityTypes.AddOwner)
                    End If

                    ' when reg is saved, the invalid activity is also marked as processed
                    oReg.Save()

                ElseIf nInvalidRegActivityCount > 0 Then
                    oReg.Save()
                End If
            ElseIf nInvalidRegActivityCount > 0 Then
                oReg.Save()
            End If

            CheckForRegistrationActivity()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

#Region "Envelopes and Labels"
    Private Sub btnEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnvelopes.Click
        Dim arrAddress(4) As String
        Try

            arrAddress(0) = pOwn.Addresses.AddressLine1
            arrAddress(1) = pOwn.Addresses.AddressLine2
            arrAddress(2) = pOwn.Addresses.City
            arrAddress(3) = pOwn.Addresses.State
            arrAddress(4) = pOwn.Addresses.Zip

            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(Me.txtOwnerName.Text, arrAddress, "REG", pOwn.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Addresses.AddressLine1
            arrAddress(1) = pOwn.Addresses.AddressLine2
            arrAddress(2) = pOwn.Addresses.City
            arrAddress(3) = pOwn.Addresses.State
            arrAddress(4) = pOwn.Addresses.Zip
            'strAddress = pOwn.Addresses.AddressLine1 + "," + pOwn.Addresses.AddressLine2 + "," + pOwn.Addresses.City + "," + pOwn.Addresses.State + "," + pOwn.Addresses.Zip
            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtOwnerName.Text, arrAddress, "REG", pOwn.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityEnvelopes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacilityEnvelopes.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(Me.txtFacilityName.Text, arrAddress, "REG", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtFacilityName.Text, arrAddress, "REG", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region


    Private Sub chkProhibition_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkProhibition.CheckedChanged
        'This procedure is not used
        If bolLoading Then Exit Sub

        Dim LocalUserSettings As Microsoft.Win32.Registry
        Dim conSQLConnection As New SqlConnection
        Dim cmdSQLCommand As New SqlCommand
        conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")

        conSQLConnection.Open()
        cmdSQLCommand.Connection = conSQLConnection

        If (Me.chkProhibition.Checked) Then
            cmdSQLCommand.CommandText = "Insert into tblReg_Prohibition values(" + nFacilityID.ToString + ",'" + Today().ToString + "')"
            cmdSQLCommand.ExecuteNonQuery()
            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()
            MsgBox("Facility " + nFacilityID.ToString + " has been added to the Delivery Prohibition List.")
        Else
            cmdSQLCommand.CommandText = "delete from tblReg_Prohibition where Facility_id = " + nFacilityID.ToString
            cmdSQLCommand.ExecuteNonQuery()
            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()
            MsgBox("Facility " + nFacilityID.ToString + " has been removed from the Delivery Prohibition List.")
        End If
    End Sub

    Private Sub chkTankProhibition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTankProhibition.CheckedChanged
        If bolLoading Then Exit Sub
        If autoChange Then Exit Sub
        Dim LocalUserSettings As Microsoft.Win32.Registry
        Dim conSQLConnection As New SqlConnection
        Dim cmdSQLCommand As New SqlCommand
        ' Dim cmdSQLReader As New SqlCommand
        '  Dim prohibitionReader As SqlDataReader
        ' Dim hasRecord As Boolean
        ' hasRecord = False
        conSQLConnection.ConnectionString = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")

        conSQLConnection.Open()
        cmdSQLCommand.Connection = conSQLConnection
        ' cmdSQLReader.Connection = conSQLConnection

        If (Me.chkTankProhibition.Checked) Then
            ' cmdSQLCommand.CommandText = "select * from tblReg_prohibition where facility_id = " + nFacilityID.ToString + " and tank_id = " + nTankID.ToString
            ' prohibitionReader = cmdSQLCommand.ExecuteReader()
            '  If Not prohibitionReader.HasRows() Then
            '  hasRecord = True
            ' Else
            '   hasRecord = False
            ' End If
            ' prohibitionReader.Close()
            '  cmdSQLCommand.Dispose()
            '   cmdSQLCommand.Connection = conSQLConnection
            '  If Not hasRecord Then
            cmdSQLCommand.CommandText = "If not exists (select * from tblReg_prohibition where facility_id = " + nFacilityID.ToString + " and tank_id = " + nTankID.ToString + ") Insert into tblReg_Prohibition values(" + nFacilityID.ToString + "," + nTankID.ToString + ",'" + Today().ToString + "')"
            cmdSQLCommand.ExecuteNonQuery()
            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()
            MsgBox("Facility " + nFacilityID.ToString + "-Tank ID " + nTankID.ToString + " has been added to the Delivery Prohibition List.")
            autoChange = False
            ' End If
        Else
            cmdSQLCommand.CommandText = "delete from tblReg_Prohibition where Facility_id = " + nFacilityID.ToString + " and Tank_id = " + nTankID.ToString
            cmdSQLCommand.ExecuteNonQuery()
            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()
            autoChange = False
        End If
        conSQLConnection.Close()
        cmdSQLCommand.Dispose()
        ' cmdSQLReader.Dispose()
        conSQLConnection.Dispose()

    End Sub


    Private Sub dtPickAssess_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickAssess.ValueChanged
        Me.btnFacilitySave.Enabled = True
        Me.btnFacilityCancel.Enabled = True
    End Sub

    Private Sub dtPickAssess_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtPickAssess.MouseDown
        Me.dtPickAssess.Format = DateTimePickerFormat.Short
        If dtPickAssess.Checked = False Then
            UIUtilsGen.SetDatePickerValue(dtPickAssess, dtNullDate)
            If Me.btnFacilitySave.Enabled = False Then
                Me.btnFacilitySave.Enabled = True
                Me.btnFacilityCancel.Enabled = True
            End If
        End If
    End Sub

    Private Sub dtPickSpillPreventionInstalled_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtPickSpillPreventionInstalled.MouseDown
        Me.dtPickSpillPreventionInstalled.Format = DateTimePickerFormat.Short
        If dtPickSpillPreventionInstalled.Checked = False Then
            UIUtilsGen.SetDatePickerValue(dtPickSpillPreventionInstalled, dtNullDate)
        End If
    End Sub
    Private Sub dtPickSpillPreventionLastTested_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtPickSpillPreventionLastTested.MouseDown
        Me.dtPickSpillPreventionLastTested.Format = DateTimePickerFormat.Short
        If Me.dtPickSpillPreventionLastTested.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtPickSpillPreventionLastTested, dtNullDate)
        End If
    End Sub
    Private Sub dtPickOverfillPreventionLastInspected_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtPickOverfillPreventionLastInspected.MouseDown
        Me.dtPickOverfillPreventionLastInspected.Format = DateTimePickerFormat.Short
        If Me.dtPickOverfillPreventionLastInspected.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtPickOverfillPreventionLastInspected, dtNullDate)
        End If
    End Sub
    Private Sub dtPickSecondaryContainmentLastInspected_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtPickSecondaryContainmentLastInspected.MouseDown
        Me.dtPickSecondaryContainmentLastInspected.Format = DateTimePickerFormat.Short
        If Me.dtPickSecondaryContainmentLastInspected.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtPickSecondaryContainmentLastInspected, dtNullDate)
        End If
    End Sub
    Private Sub dtPickElectronicDeviceInspected_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtPickElectronicDeviceInspected.MouseDown
        Me.dtPickElectronicDeviceInspected.Format = DateTimePickerFormat.Short
        If Me.dtPickElectronicDeviceInspected.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtPickElectronicDeviceInspected, dtNullDate)
        End If
    End Sub
    Private Sub dtPickATGLastInspected_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtPickATGLastInspected.MouseDown
        Me.dtPickATGLastInspected.Format = DateTimePickerFormat.Short
        If Me.dtPickATGLastInspected.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtPickATGLastInspected, dtNullDate)
        End If
    End Sub
    Private Sub dtPickOverfillPreventionInstalled_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtPickOverfillPreventionInstalled.MouseDown
        Me.dtPickOverfillPreventionInstalled.Format = DateTimePickerFormat.Short
        If Me.dtPickOverfillPreventionInstalled.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtPickOverfillPreventionInstalled, dtNullDate)
        End If
    End Sub
    Private Sub dtSheerValueTest_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtSheerValueTest.MouseDown
        Me.dtSheerValueTest.Format = DateTimePickerFormat.Short
        If Me.dtSheerValueTest.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtSheerValueTest, dtNullDate)
        End If
    End Sub
    Private Sub dtSecondaryContainmentInspected_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtSecondaryContainmentInspected.MouseDown
        Me.dtSecondaryContainmentInspected.Format = DateTimePickerFormat.Short
        If Me.dtSecondaryContainmentInspected.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtSecondaryContainmentInspected, dtNullDate)
        End If
    End Sub
    Private Sub dtElectronicDeviceInspected_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dtElectronicDeviceInspected.MouseDown
        Me.dtElectronicDeviceInspected.Format = DateTimePickerFormat.Short
        If Me.dtElectronicDeviceInspected.Checked = False Then
            UIUtilsGen.SetDatePickerValue(Me.dtElectronicDeviceInspected, dtNullDate)
        End If
    End Sub

End Class