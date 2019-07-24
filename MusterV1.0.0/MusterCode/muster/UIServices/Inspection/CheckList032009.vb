Public Class CheckList
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private WithEvents oInspection As MUSTER.BusinessLogic.pInspection
    Private WithEvents AddressForm As Address
    Private WithEvents oTank As MUSTER.BusinessLogic.pTank
    Private WithEvents oPipe As MUSTER.BusinessLogic.pPipe
    Private WithEvents ltrGen As MUSTER.BusinessLogic.pLetterGen
    Private WithEvents pFlag As MUSTER.BusinessLogic.pFlag
    Private WithEvents SF As ShowFlags
    Private frmCLProgress As CheckListProgress
    Private bolLoading As Boolean = False
    Private bolReadOnly, bolPrint, bolCancel, bolSaveDataToArchiveTbls As Boolean
    Private strAddress, strState, strCity, strCounty, strFIPS, strZipCode, mode As String
    Dim dtNullDate As Date = CDate("01/01/0001")
    Private regenerateCheckListItems As Boolean = False
    Private tankpipetermcellbeingupdated As Boolean = False
    Private tankCellUpdated As Boolean = False
    Private tankDateCellUpdated As Boolean = False
    Private tankSizeUpdated As Boolean = False
    Private pipeCellUpdated As Boolean = False
    Private pipeDateCellUpdated As Boolean = False
    Private termCellUpdated As Boolean = False
    Private termDateCellUpdated As Boolean = False
    Private bolCheckListPrinting As Boolean = False
    Private bolErrOccurred As Boolean = False
    Private bolFromBtnTanksPipes As Boolean = False
    Private hasCPTank, hasCPPipe, hasCPTerm As Boolean
    Dim btnPrevSelected As Button = Nothing
    ' valuelists to maintain data
    Dim vListCompSubstance, vListMaterialTank, vListMaterialPipe, vListOverFill, vListTypeTermDisp, vListTypeTermTank, vListTDispTermCP, vListTankTermCP, vListTankNum, vListPipeNum, vListTermNum, vListTime, vListType, vListBrand, vListCPType, vListInspectors As Infragistics.Win.ValueList
    Dim dv As DataView
    ' see class definition below
    Dim rp As New Remove_Pencil
    Dim returnVal As String = String.Empty
    ' enum to control which grid in TanksPipes panel needs to be refreshed / recalculated
    Private Enum TankPipeTermGrid
        Tank
        Pipe
        Term
        All
    End Enum
    Private Enum TankPipeTermCPReading
        Tank
        Pipe
        Term
    End Enum
    ' Contact Mgmt variables
    Private WithEvents ContactFrm As Contacts
    Private WithEvents objCntSearch As ContactSearch
    Dim dsContacts As DataSet
    Dim strFilterString As String = String.Empty
    Dim strFacilityIdTags As String
    Dim result As DialogResult
    Friend CallingForm As Form
    Private moduleID As Integer
    Private bol354Hidden, bol363Hidden, bol376Hidden As Boolean

    ' variables for SetupTankPipeTerm
    ' Tank
    'Dim CPTypeCount, CPInstalledCount, CPTestedCount, LeakDetectionCount, OverFillCount, LastUsedCount, _
    'LinedCount, LiningInspectedCount, PTTCount As Integer
    ' Pipe
    'Dim BrandCount, PriLeakDetectionCount, SecLeakDetectionCount, ALLDTestedCount, PipeLastUsedCount As Integer
    ' Term
    'Dim TermCPTestedCount, TermDispCPTypeCount, TermTankCPTypeCount As Integer

#End Region
#Region "Windows Form Designer generated code "

    Public Sub New(ByRef oInspec As MUSTER.BusinessLogic.pInspection, Optional ByVal [readOnly] As Boolean = False, Optional ByVal print As Boolean = False, Optional ByVal fromModule As Integer = UIUtilsGen.ModuleID.Inspection)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        Cursor.Current = Cursors.AppStarting
        bolCancel = False
        bolSaveDataToArchiveTbls = False
        oInspection = oInspec
        pnlMaster.Dock = DockStyle.Fill
        pnlTanksPipes.Dock = DockStyle.Fill
        pnlInspectionCitations.Dock = DockStyle.Fill
        pnlReg.Dock = DockStyle.Fill
        pnlSpill.Dock = DockStyle.Fill
        pnlCP.Dock = DockStyle.Fill
        pnlTankLeak.Dock = DockStyle.Fill
        pnlPipeLeak.Dock = DockStyle.Fill
        pnlCatLeak.Dock = DockStyle.Fill
        pnlVisual.Dock = DockStyle.Fill
        pnlTOS.Dock = DockStyle.Fill
        pnlComments.Dock = DockStyle.Fill
        pnlMW.Dock = DockStyle.Fill
        pnlSOC.Dock = DockStyle.Fill
        bolReadOnly = [readOnly]
        bolPrint = print
        moduleID = fromModule
        oInspection.RetrieveCheckListInfo(oInspection.ID, oInspection.FacilityID, oInspection.OwnerID, [readOnly])
        txtAddress.Tag = oInspection.CheckListMaster.Owner.Facilities.AddressID
        Me.Text = IIf([readOnly], "View", "Edit") + " CheckList for Facility : " + oInspection.CheckListMaster.Owner.Facilities.Name + " (" + oInspection.CheckListMaster.Owner.Facilities.ID.ToString + ")"
        ToggleButtonAppearance(btnMaster)

        LoadChecklistMaster()
        If bolReadOnly Then
            MakeFormReadOnly(bolReadOnly)
            If Date.Compare(oInspection.Completed, dtNullDate) = 0 Then
                If oInspection.ScheduledBy.ToUpper = MusterContainer.AppUser.ID.ToUpper Then
                    btnUnsubmit.Enabled = True
                End If
            End If
        Else
            If Date.Compare(oInspection.SubmittedDate, dtNullDate) = 0 Then
                btnSubmitToCE.Enabled = True
                If Date.Compare(oInspection.Completed, dtNullDate) = 0 Then
                    btnUnsubmit.Enabled = True
                Else
                    btnUnsubmit.Enabled = False
                End If
            Else
                btnSubmitToCE.Enabled = False
                If Date.Compare(oInspection.Completed, dtNullDate) = 0 Then
                    btnUnsubmit.Enabled = True
                Else
                    btnUnsubmit.Enabled = False
                End If
            End If
            btnSave.Enabled = oInspection.IsDirty Or oInspection.CheckListMaster.colIsDirty
        End If
        hasCPTank = False
        hasCPPipe = False
        hasCPTerm = False
        If MusterContainer.AppUser.HEAD_INSPECTION Or Date.Compare(oInspection.Completed, dtNullDate) <> 0 Then
            btnSoc.Visible = True
        Else
            btnSoc.Visible = False
        End If
        Cursor.Current = Cursors.Default
        LoadContacts(ugContacts, oInspection.OwnerID, 9)

        pFlag = New MUSTER.BusinessLogic.pFlag
        LoadBarometer()
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
    Friend WithEvents pnlChecklistBottom As System.Windows.Forms.Panel
    Friend WithEvents btnUnsubmit As System.Windows.Forms.Button
    Friend WithEvents btnSubmitToCE As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents pnlChecklistControls As System.Windows.Forms.Panel
    Friend WithEvents btnSoc As System.Windows.Forms.Button
    Friend WithEvents btnComments As System.Windows.Forms.Button
    Friend WithEvents btnTos As System.Windows.Forms.Button
    Friend WithEvents btnVisual As System.Windows.Forms.Button
    Friend WithEvents btnCatLeak As System.Windows.Forms.Button
    Friend WithEvents btnPipeLeak As System.Windows.Forms.Button
    Friend WithEvents btnTankLeak As System.Windows.Forms.Button
    Friend WithEvents btnCp As System.Windows.Forms.Button
    Friend WithEvents btnSpill As System.Windows.Forms.Button
    Friend WithEvents btnReg As System.Windows.Forms.Button
    Friend WithEvents btnCitations As System.Windows.Forms.Button
    Friend WithEvents btnMaster As System.Windows.Forms.Button
    Friend WithEvents btnTanksPipes As System.Windows.Forms.Button
    Friend WithEvents pnlChecklistDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlMaster As System.Windows.Forms.Panel
    Friend WithEvents pnlMasterDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlMasterDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblMasterDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlTanksPipes As System.Windows.Forms.Panel
    Friend WithEvents pnlTanksPipesDetails As System.Windows.Forms.Panel
    Friend WithEvents ugTerminations As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblTerminations As System.Windows.Forms.Label
    Friend WithEvents ugPipes As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblPipes As System.Windows.Forms.Label
    Friend WithEvents ugTanks As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblTanks As System.Windows.Forms.Label
    Friend WithEvents pnlTanksPipesDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblTanksPipesDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlInspectionCitations As System.Windows.Forms.Panel
    Friend WithEvents ugCitation As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlInspecCitationsDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblInspecCitationsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlReg As System.Windows.Forms.Panel
    Friend WithEvents pnlRegDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblReg As System.Windows.Forms.Label
    Friend WithEvents pnlTanksTop As System.Windows.Forms.Panel
    Friend WithEvents pnlTanksDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlPipeTop As System.Windows.Forms.Panel
    Friend WithEvents pnlPipeDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlTerminationsTop As System.Windows.Forms.Panel
    Friend WithEvents pnlTerminationsDetails As System.Windows.Forms.Panel
    Friend WithEvents ugReg As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlSpill As System.Windows.Forms.Panel
    Friend WithEvents pnlSpillDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblSpill As System.Windows.Forms.Label
    Friend WithEvents ugSpill As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlCP As System.Windows.Forms.Panel
    Friend WithEvents pnlTankLeak As System.Windows.Forms.Panel
    Friend WithEvents pnlTankLeakDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblTankLeak As System.Windows.Forms.Label
    Friend WithEvents ugTankLeak As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlPipeLeak As System.Windows.Forms.Panel
    Friend WithEvents pnlPipeLeakDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblPipeLeak As System.Windows.Forms.Label
    Friend WithEvents ugPipeLeak As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlCatLeak As System.Windows.Forms.Panel
    Friend WithEvents pnlCatLeakDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblCatLeak As System.Windows.Forms.Label
    Friend WithEvents ugCatLeak As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlVisual As System.Windows.Forms.Panel
    Friend WithEvents pnlVisualDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblVisual As System.Windows.Forms.Label
    Friend WithEvents ugVisual As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlTOS As System.Windows.Forms.Panel
    Friend WithEvents pnlComments As System.Windows.Forms.Panel
    Friend WithEvents pnlSOC As System.Windows.Forms.Panel
    Friend WithEvents pnlTOSDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblTOS As System.Windows.Forms.Label
    Friend WithEvents pnlSOCDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblSOC As System.Windows.Forms.Label
    Friend WithEvents ugTOS As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlCPDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblCP As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents pnlCommentsDisplay As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents chkSOCSketchAttached As System.Windows.Forms.CheckBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ugSOC As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugCP As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlCPAdd As System.Windows.Forms.Panel
    Friend WithEvents btnAddTankCP As System.Windows.Forms.Button
    Friend WithEvents btnAddPipeCP As System.Windows.Forms.Button
    Friend WithEvents btnAddTermCP As System.Windows.Forms.Button
    Friend WithEvents pnlMasterDetailsTop As System.Windows.Forms.Panel
    Friend WithEvents txtOwner As System.Windows.Forms.TextBox
    Friend WithEvents lblOwner As System.Windows.Forms.Label
    Friend WithEvents txtFacility As System.Windows.Forms.TextBox
    Friend WithEvents lblFacility As System.Windows.Forms.Label
    Friend WithEvents chkCapCandidate As System.Windows.Forms.CheckBox
    Friend WithEvents lblOwnersPhone As System.Windows.Forms.Label
    Friend WithEvents txtOwnersPhone As System.Windows.Forms.TextBox
    Public WithEvents mskTxtOwnerPhone As AxMSMask.AxMaskEdBox
    Friend WithEvents lblZip As System.Windows.Forms.Label
    Friend WithEvents lblState As System.Windows.Forms.Label
    Friend WithEvents txtOwnerZip As System.Windows.Forms.TextBox
    Friend WithEvents txtOwnerState As System.Windows.Forms.TextBox
    Friend WithEvents txtOwnerCity As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerCity As System.Windows.Forms.Label
    Friend WithEvents txtOwnersAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnersAddress As System.Windows.Forms.Label
    Friend WithEvents txtOwnersRep As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnersRap As System.Windows.Forms.Label
    Friend WithEvents lblLustOwner As System.Windows.Forms.Label
    Friend WithEvents txtUstOwner As System.Windows.Forms.TextBox
    Friend WithEvents lblLongitudeSec As System.Windows.Forms.Label
    Friend WithEvents lblLongitudeMin As System.Windows.Forms.Label
    Friend WithEvents txtLongitudeSec As System.Windows.Forms.TextBox
    Friend WithEvents txtLongitudeMin As System.Windows.Forms.TextBox
    Friend WithEvents lblLongitudeDegree As System.Windows.Forms.Label
    Friend WithEvents txtLongitudeDegree As System.Windows.Forms.TextBox
    Friend WithEvents lblLogitudeW As System.Windows.Forms.Label
    Friend WithEvents lblLatitudeSec As System.Windows.Forms.Label
    Friend WithEvents txtLatitudeSec As System.Windows.Forms.TextBox
    Friend WithEvents lblLatitudeMin As System.Windows.Forms.Label
    Friend WithEvents txtLatitudeMin As System.Windows.Forms.TextBox
    Friend WithEvents lblLatitudeDegree As System.Windows.Forms.Label
    Friend WithEvents txtLatitudeDegree As System.Windows.Forms.TextBox
    Friend WithEvents lblLatitudeN As System.Windows.Forms.Label
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents pnlMasterDetailsGrid As System.Windows.Forms.Panel
    Friend WithEvents ugInspector As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlContact As System.Windows.Forms.Panel
    Friend WithEvents chkOwnerShowActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents btnOwnerModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAddSearchContact As System.Windows.Forms.Button
    Friend WithEvents pnlContactTop As System.Windows.Forms.Panel
    Friend WithEvents pnlContactBottom As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ugContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblContact As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnIndvLicFlag As System.Windows.Forms.Button
    Friend WithEvents btnInsFlag As System.Windows.Forms.Button
    Friend WithEvents btnCandEFlag As System.Windows.Forms.Button
    Friend WithEvents btnFinFlag As System.Windows.Forms.Button
    Friend WithEvents btnLUSFlag As System.Windows.Forms.Button
    Friend WithEvents btnFirmLicFlag As System.Windows.Forms.Button
    Friend WithEvents btnCloFlag As System.Windows.Forms.Button
    Friend WithEvents btnFeeFlag As System.Windows.Forms.Button
    Friend WithEvents btnFacFlag As System.Windows.Forms.Button
    Friend WithEvents btnOwnerFlag As System.Windows.Forms.Button
    Friend WithEvents btnFlags As System.Windows.Forms.Button
    Friend WithEvents pnlTankLeakAdd As System.Windows.Forms.Panel
    Friend WithEvents pnlPipeLeakAdd As System.Windows.Forms.Panel
    Friend WithEvents btnAddTankMW As System.Windows.Forms.Button
    Friend WithEvents btnAddPipeMW As System.Windows.Forms.Button
    Friend WithEvents lblFacilityName As System.Windows.Forms.Label
    Friend WithEvents txtFacilityName As System.Windows.Forms.TextBox
    Friend WithEvents btnMW As System.Windows.Forms.Button
    Friend WithEvents pnlMW As System.Windows.Forms.Panel
    Friend WithEvents ugMW As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlMWDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblMW As System.Windows.Forms.Label
    Friend WithEvents pnlMWAdd As System.Windows.Forms.Panel
    Friend WithEvents btnMWAdd As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents lblLustActive As System.Windows.Forms.Label
    Friend WithEvents lblLustActiveValue As System.Windows.Forms.Label
    Friend WithEvents lblLustPM As System.Windows.Forms.Label
    Friend WithEvents lblLustPMValue As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CheckList))
        Me.pnlChecklistBottom = New System.Windows.Forms.Panel
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
        Me.btnUnsubmit = New System.Windows.Forms.Button
        Me.btnSubmitToCE = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFlags = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.pnlChecklistControls = New System.Windows.Forms.Panel
        Me.btnSoc = New System.Windows.Forms.Button
        Me.btnComments = New System.Windows.Forms.Button
        Me.btnTos = New System.Windows.Forms.Button
        Me.btnVisual = New System.Windows.Forms.Button
        Me.btnCatLeak = New System.Windows.Forms.Button
        Me.btnPipeLeak = New System.Windows.Forms.Button
        Me.btnTankLeak = New System.Windows.Forms.Button
        Me.btnCp = New System.Windows.Forms.Button
        Me.btnSpill = New System.Windows.Forms.Button
        Me.btnReg = New System.Windows.Forms.Button
        Me.btnCitations = New System.Windows.Forms.Button
        Me.btnMaster = New System.Windows.Forms.Button
        Me.btnTanksPipes = New System.Windows.Forms.Button
        Me.btnMW = New System.Windows.Forms.Button
        Me.pnlChecklistDetails = New System.Windows.Forms.Panel
        Me.pnlMW = New System.Windows.Forms.Panel
        Me.ugMW = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlMWAdd = New System.Windows.Forms.Panel
        Me.btnMWAdd = New System.Windows.Forms.Button
        Me.pnlMWDisplay = New System.Windows.Forms.Panel
        Me.lblMW = New System.Windows.Forms.Label
        Me.pnlMaster = New System.Windows.Forms.Panel
        Me.pnlMasterDetails = New System.Windows.Forms.Panel
        Me.pnlContact = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.ugContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlContactBottom = New System.Windows.Forms.Panel
        Me.btnOwnerModifyContact = New System.Windows.Forms.Button
        Me.btnOwnerDeleteContact = New System.Windows.Forms.Button
        Me.btnOwnerAssociateContact = New System.Windows.Forms.Button
        Me.btnOwnerAddSearchContact = New System.Windows.Forms.Button
        Me.pnlContactTop = New System.Windows.Forms.Panel
        Me.lblContact = New System.Windows.Forms.Label
        Me.chkOwnerShowContactsforAllModules = New System.Windows.Forms.CheckBox
        Me.chkOwnerShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkOwnerShowActiveOnly = New System.Windows.Forms.CheckBox
        Me.pnlMasterDetailsGrid = New System.Windows.Forms.Panel
        Me.ugInspector = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlMasterDetailsTop = New System.Windows.Forms.Panel
        Me.txtOwner = New System.Windows.Forms.TextBox
        Me.lblOwner = New System.Windows.Forms.Label
        Me.txtFacility = New System.Windows.Forms.TextBox
        Me.lblFacility = New System.Windows.Forms.Label
        Me.lblOwnersPhone = New System.Windows.Forms.Label
        Me.mskTxtOwnerPhone = New AxMSMask.AxMaskEdBox
        Me.txtOwnersPhone = New System.Windows.Forms.TextBox
        Me.txtOwnerZip = New System.Windows.Forms.TextBox
        Me.txtOwnerState = New System.Windows.Forms.TextBox
        Me.txtOwnerCity = New System.Windows.Forms.TextBox
        Me.lblOwnerCity = New System.Windows.Forms.Label
        Me.txtOwnersAddress = New System.Windows.Forms.TextBox
        Me.lblOwnersAddress = New System.Windows.Forms.Label
        Me.txtOwnersRep = New System.Windows.Forms.TextBox
        Me.lblOwnersRap = New System.Windows.Forms.Label
        Me.lblLustOwner = New System.Windows.Forms.Label
        Me.txtUstOwner = New System.Windows.Forms.TextBox
        Me.lblLongitudeSec = New System.Windows.Forms.Label
        Me.lblLongitudeMin = New System.Windows.Forms.Label
        Me.txtLongitudeSec = New System.Windows.Forms.TextBox
        Me.txtLongitudeMin = New System.Windows.Forms.TextBox
        Me.lblLongitudeDegree = New System.Windows.Forms.Label
        Me.txtLongitudeDegree = New System.Windows.Forms.TextBox
        Me.lblLogitudeW = New System.Windows.Forms.Label
        Me.lblLatitudeSec = New System.Windows.Forms.Label
        Me.txtLatitudeSec = New System.Windows.Forms.TextBox
        Me.lblLatitudeMin = New System.Windows.Forms.Label
        Me.txtLatitudeMin = New System.Windows.Forms.TextBox
        Me.txtLatitudeDegree = New System.Windows.Forms.TextBox
        Me.lblLatitudeN = New System.Windows.Forms.Label
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.lblAddress = New System.Windows.Forms.Label
        Me.lblLatitudeDegree = New System.Windows.Forms.Label
        Me.chkCapCandidate = New System.Windows.Forms.CheckBox
        Me.lblState = New System.Windows.Forms.Label
        Me.lblZip = New System.Windows.Forms.Label
        Me.lblFacilityName = New System.Windows.Forms.Label
        Me.txtFacilityName = New System.Windows.Forms.TextBox
        Me.lblLustActive = New System.Windows.Forms.Label
        Me.lblLustActiveValue = New System.Windows.Forms.Label
        Me.lblLustPM = New System.Windows.Forms.Label
        Me.lblLustPMValue = New System.Windows.Forms.Label
        Me.pnlMasterDisplay = New System.Windows.Forms.Panel
        Me.lblMasterDisplay = New System.Windows.Forms.Label
        Me.pnlComments = New System.Windows.Forms.Panel
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.pnlCommentsDisplay = New System.Windows.Forms.Panel
        Me.lblComments = New System.Windows.Forms.Label
        Me.pnlTOS = New System.Windows.Forms.Panel
        Me.ugTOS = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTOSDisplay = New System.Windows.Forms.Panel
        Me.lblTOS = New System.Windows.Forms.Label
        Me.pnlVisual = New System.Windows.Forms.Panel
        Me.ugVisual = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlVisualDisplay = New System.Windows.Forms.Panel
        Me.lblVisual = New System.Windows.Forms.Label
        Me.pnlCatLeak = New System.Windows.Forms.Panel
        Me.ugCatLeak = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlCatLeakDisplay = New System.Windows.Forms.Panel
        Me.lblCatLeak = New System.Windows.Forms.Label
        Me.pnlPipeLeak = New System.Windows.Forms.Panel
        Me.ugPipeLeak = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlPipeLeakAdd = New System.Windows.Forms.Panel
        Me.btnAddPipeMW = New System.Windows.Forms.Button
        Me.pnlPipeLeakDisplay = New System.Windows.Forms.Panel
        Me.lblPipeLeak = New System.Windows.Forms.Label
        Me.pnlTankLeak = New System.Windows.Forms.Panel
        Me.ugTankLeak = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTankLeakAdd = New System.Windows.Forms.Panel
        Me.btnAddTankMW = New System.Windows.Forms.Button
        Me.pnlTankLeakDisplay = New System.Windows.Forms.Panel
        Me.lblTankLeak = New System.Windows.Forms.Label
        Me.pnlCP = New System.Windows.Forms.Panel
        Me.ugCP = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlCPAdd = New System.Windows.Forms.Panel
        Me.btnAddTankCP = New System.Windows.Forms.Button
        Me.btnAddPipeCP = New System.Windows.Forms.Button
        Me.btnAddTermCP = New System.Windows.Forms.Button
        Me.pnlCPDisplay = New System.Windows.Forms.Panel
        Me.lblCP = New System.Windows.Forms.Label
        Me.pnlSpill = New System.Windows.Forms.Panel
        Me.ugSpill = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlSpillDisplay = New System.Windows.Forms.Panel
        Me.lblSpill = New System.Windows.Forms.Label
        Me.pnlReg = New System.Windows.Forms.Panel
        Me.pnlRegDisplay = New System.Windows.Forms.Panel
        Me.lblReg = New System.Windows.Forms.Label
        Me.ugReg = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlInspectionCitations = New System.Windows.Forms.Panel
        Me.pnlInspecCitationsDisplay = New System.Windows.Forms.Panel
        Me.lblInspecCitationsDisplay = New System.Windows.Forms.Label
        Me.ugCitation = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTanksPipes = New System.Windows.Forms.Panel
        Me.pnlTanksPipesDetails = New System.Windows.Forms.Panel
        Me.pnlTerminationsDetails = New System.Windows.Forms.Panel
        Me.ugTerminations = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTerminationsTop = New System.Windows.Forms.Panel
        Me.lblTerminations = New System.Windows.Forms.Label
        Me.pnlPipeDetails = New System.Windows.Forms.Panel
        Me.ugPipes = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlPipeTop = New System.Windows.Forms.Panel
        Me.lblPipes = New System.Windows.Forms.Label
        Me.pnlTanksDetails = New System.Windows.Forms.Panel
        Me.ugTanks = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTanksTop = New System.Windows.Forms.Panel
        Me.lblTanks = New System.Windows.Forms.Label
        Me.pnlTanksPipesDisplay = New System.Windows.Forms.Panel
        Me.lblTanksPipesDisplay = New System.Windows.Forms.Label
        Me.pnlSOC = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.ugSOC = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.chkSOCSketchAttached = New System.Windows.Forms.CheckBox
        Me.pnlSOCDisplay = New System.Windows.Forms.Panel
        Me.lblSOC = New System.Windows.Forms.Label
        Me.pnlChecklistBottom.SuspendLayout()
        Me.pnlChecklistControls.SuspendLayout()
        Me.pnlChecklistDetails.SuspendLayout()
        Me.pnlMW.SuspendLayout()
        CType(Me.ugMW, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMWAdd.SuspendLayout()
        Me.pnlMWDisplay.SuspendLayout()
        Me.pnlMaster.SuspendLayout()
        Me.pnlMasterDetails.SuspendLayout()
        Me.pnlContact.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.ugContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlContactBottom.SuspendLayout()
        Me.pnlContactTop.SuspendLayout()
        Me.pnlMasterDetailsGrid.SuspendLayout()
        CType(Me.ugInspector, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMasterDetailsTop.SuspendLayout()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMasterDisplay.SuspendLayout()
        Me.pnlComments.SuspendLayout()
        Me.pnlCommentsDisplay.SuspendLayout()
        Me.pnlTOS.SuspendLayout()
        CType(Me.ugTOS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTOSDisplay.SuspendLayout()
        Me.pnlVisual.SuspendLayout()
        CType(Me.ugVisual, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlVisualDisplay.SuspendLayout()
        Me.pnlCatLeak.SuspendLayout()
        CType(Me.ugCatLeak, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCatLeakDisplay.SuspendLayout()
        Me.pnlPipeLeak.SuspendLayout()
        CType(Me.ugPipeLeak, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPipeLeakAdd.SuspendLayout()
        Me.pnlPipeLeakDisplay.SuspendLayout()
        Me.pnlTankLeak.SuspendLayout()
        CType(Me.ugTankLeak, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTankLeakAdd.SuspendLayout()
        Me.pnlTankLeakDisplay.SuspendLayout()
        Me.pnlCP.SuspendLayout()
        CType(Me.ugCP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCPAdd.SuspendLayout()
        Me.pnlCPDisplay.SuspendLayout()
        Me.pnlSpill.SuspendLayout()
        CType(Me.ugSpill, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlSpillDisplay.SuspendLayout()
        Me.pnlReg.SuspendLayout()
        Me.pnlRegDisplay.SuspendLayout()
        CType(Me.ugReg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlInspectionCitations.SuspendLayout()
        Me.pnlInspecCitationsDisplay.SuspendLayout()
        CType(Me.ugCitation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTanksPipes.SuspendLayout()
        Me.pnlTanksPipesDetails.SuspendLayout()
        Me.pnlTerminationsDetails.SuspendLayout()
        CType(Me.ugTerminations, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTerminationsTop.SuspendLayout()
        Me.pnlPipeDetails.SuspendLayout()
        CType(Me.ugPipes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPipeTop.SuspendLayout()
        Me.pnlTanksDetails.SuspendLayout()
        CType(Me.ugTanks, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTanksTop.SuspendLayout()
        Me.pnlTanksPipesDisplay.SuspendLayout()
        Me.pnlSOC.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.ugSOC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.pnlSOCDisplay.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlChecklistBottom
        '
        Me.pnlChecklistBottom.Controls.Add(Me.btnIndvLicFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnInsFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnCandEFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnFinFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnLUSFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnFirmLicFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnCloFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnFeeFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnFacFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnOwnerFlag)
        Me.pnlChecklistBottom.Controls.Add(Me.btnUnsubmit)
        Me.pnlChecklistBottom.Controls.Add(Me.btnSubmitToCE)
        Me.pnlChecklistBottom.Controls.Add(Me.btnCancel)
        Me.pnlChecklistBottom.Controls.Add(Me.btnSave)
        Me.pnlChecklistBottom.Controls.Add(Me.btnClose)
        Me.pnlChecklistBottom.Controls.Add(Me.btnFlags)
        Me.pnlChecklistBottom.Controls.Add(Me.btnPrint)
        Me.pnlChecklistBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlChecklistBottom.Location = New System.Drawing.Point(0, 669)
        Me.pnlChecklistBottom.Name = "pnlChecklistBottom"
        Me.pnlChecklistBottom.Size = New System.Drawing.Size(1028, 40)
        Me.pnlChecklistBottom.TabIndex = 37
        '
        'btnIndvLicFlag
        '
        Me.btnIndvLicFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnIndvLicFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnIndvLicFlag.Location = New System.Drawing.Point(976, 8)
        Me.btnIndvLicFlag.Name = "btnIndvLicFlag"
        Me.btnIndvLicFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnIndvLicFlag.TabIndex = 51
        Me.btnIndvLicFlag.Text = "Indv"
        Me.btnIndvLicFlag.Visible = False
        '
        'btnInsFlag
        '
        Me.btnInsFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnInsFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnInsFlag.Location = New System.Drawing.Point(896, 8)
        Me.btnInsFlag.Name = "btnInsFlag"
        Me.btnInsFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnInsFlag.TabIndex = 50
        Me.btnInsFlag.Text = "Insp"
        Me.btnInsFlag.Visible = False
        '
        'btnCandEFlag
        '
        Me.btnCandEFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnCandEFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCandEFlag.Location = New System.Drawing.Point(856, 8)
        Me.btnCandEFlag.Name = "btnCandEFlag"
        Me.btnCandEFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnCandEFlag.TabIndex = 49
        Me.btnCandEFlag.Text = "C&&E"
        Me.btnCandEFlag.Visible = False
        '
        'btnFinFlag
        '
        Me.btnFinFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnFinFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFinFlag.Location = New System.Drawing.Point(816, 8)
        Me.btnFinFlag.Name = "btnFinFlag"
        Me.btnFinFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnFinFlag.TabIndex = 48
        Me.btnFinFlag.Text = "Fin"
        Me.btnFinFlag.Visible = False
        '
        'btnLUSFlag
        '
        Me.btnLUSFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnLUSFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLUSFlag.Location = New System.Drawing.Point(776, 8)
        Me.btnLUSFlag.Name = "btnLUSFlag"
        Me.btnLUSFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnLUSFlag.TabIndex = 47
        Me.btnLUSFlag.Text = "Lust"
        Me.btnLUSFlag.Visible = False
        '
        'btnFirmLicFlag
        '
        Me.btnFirmLicFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnFirmLicFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFirmLicFlag.Location = New System.Drawing.Point(936, 8)
        Me.btnFirmLicFlag.Name = "btnFirmLicFlag"
        Me.btnFirmLicFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnFirmLicFlag.TabIndex = 46
        Me.btnFirmLicFlag.Text = "Firm"
        Me.btnFirmLicFlag.Visible = False
        '
        'btnCloFlag
        '
        Me.btnCloFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnCloFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnCloFlag.Location = New System.Drawing.Point(736, 8)
        Me.btnCloFlag.Name = "btnCloFlag"
        Me.btnCloFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnCloFlag.TabIndex = 45
        Me.btnCloFlag.Text = "Clos"
        Me.btnCloFlag.Visible = False
        '
        'btnFeeFlag
        '
        Me.btnFeeFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnFeeFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFeeFlag.Location = New System.Drawing.Point(696, 8)
        Me.btnFeeFlag.Name = "btnFeeFlag"
        Me.btnFeeFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnFeeFlag.TabIndex = 44
        Me.btnFeeFlag.Text = "Fee"
        Me.btnFeeFlag.Visible = False
        '
        'btnFacFlag
        '
        Me.btnFacFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnFacFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnFacFlag.Location = New System.Drawing.Point(656, 8)
        Me.btnFacFlag.Name = "btnFacFlag"
        Me.btnFacFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnFacFlag.TabIndex = 43
        Me.btnFacFlag.Text = "Reg"
        Me.btnFacFlag.Visible = False
        '
        'btnOwnerFlag
        '
        Me.btnOwnerFlag.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerFlag.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOwnerFlag.Location = New System.Drawing.Point(616, 8)
        Me.btnOwnerFlag.Name = "btnOwnerFlag"
        Me.btnOwnerFlag.Size = New System.Drawing.Size(40, 23)
        Me.btnOwnerFlag.TabIndex = 42
        Me.btnOwnerFlag.Text = "Own"
        Me.btnOwnerFlag.Visible = False
        '
        'btnUnsubmit
        '
        Me.btnUnsubmit.Enabled = False
        Me.btnUnsubmit.Location = New System.Drawing.Point(269, 9)
        Me.btnUnsubmit.Name = "btnUnsubmit"
        Me.btnUnsubmit.TabIndex = 41
        Me.btnUnsubmit.Text = "Unsubmit"
        '
        'btnSubmitToCE
        '
        Me.btnSubmitToCE.Location = New System.Drawing.Point(176, 9)
        Me.btnSubmitToCE.Name = "btnSubmitToCE"
        Me.btnSubmitToCE.Size = New System.Drawing.Size(88, 23)
        Me.btnSubmitToCE.TabIndex = 40
        Me.btnSubmitToCE.Text = "Submit to C&&E"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(96, 9)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 39
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(16, 9)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 38
        Me.btnSave.Text = "Save"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(429, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 39
        Me.btnClose.Text = "Close"
        '
        'btnFlags
        '
        Me.btnFlags.Location = New System.Drawing.Point(349, 8)
        Me.btnFlags.Name = "btnFlags"
        Me.btnFlags.TabIndex = 39
        Me.btnFlags.Text = "Flags"
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(509, 8)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.TabIndex = 39
        Me.btnPrint.Text = "Print"
        '
        'pnlChecklistControls
        '
        Me.pnlChecklistControls.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlChecklistControls.Controls.Add(Me.btnSoc)
        Me.pnlChecklistControls.Controls.Add(Me.btnComments)
        Me.pnlChecklistControls.Controls.Add(Me.btnTos)
        Me.pnlChecklistControls.Controls.Add(Me.btnVisual)
        Me.pnlChecklistControls.Controls.Add(Me.btnCatLeak)
        Me.pnlChecklistControls.Controls.Add(Me.btnPipeLeak)
        Me.pnlChecklistControls.Controls.Add(Me.btnTankLeak)
        Me.pnlChecklistControls.Controls.Add(Me.btnCp)
        Me.pnlChecklistControls.Controls.Add(Me.btnSpill)
        Me.pnlChecklistControls.Controls.Add(Me.btnReg)
        Me.pnlChecklistControls.Controls.Add(Me.btnCitations)
        Me.pnlChecklistControls.Controls.Add(Me.btnMaster)
        Me.pnlChecklistControls.Controls.Add(Me.btnTanksPipes)
        Me.pnlChecklistControls.Controls.Add(Me.btnMW)
        Me.pnlChecklistControls.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlChecklistControls.Location = New System.Drawing.Point(0, 0)
        Me.pnlChecklistControls.Name = "pnlChecklistControls"
        Me.pnlChecklistControls.Size = New System.Drawing.Size(104, 669)
        Me.pnlChecklistControls.TabIndex = 0
        '
        'btnSoc
        '
        Me.btnSoc.Location = New System.Drawing.Point(4, 616)
        Me.btnSoc.Name = "btnSoc"
        Me.btnSoc.Size = New System.Drawing.Size(94, 46)
        Me.btnSoc.TabIndex = 12
        Me.btnSoc.Text = "SOC"
        '
        'btnComments
        '
        Me.btnComments.Location = New System.Drawing.Point(4, 474)
        Me.btnComments.Name = "btnComments"
        Me.btnComments.Size = New System.Drawing.Size(94, 46)
        Me.btnComments.TabIndex = 10
        Me.btnComments.Text = "8 - COMMENTS"
        '
        'btnTos
        '
        Me.btnTos.Location = New System.Drawing.Point(4, 427)
        Me.btnTos.Name = "btnTos"
        Me.btnTos.Size = New System.Drawing.Size(94, 46)
        Me.btnTos.TabIndex = 9
        Me.btnTos.Text = "7 - TOS"
        '
        'btnVisual
        '
        Me.btnVisual.Location = New System.Drawing.Point(4, 381)
        Me.btnVisual.Name = "btnVisual"
        Me.btnVisual.Size = New System.Drawing.Size(94, 46)
        Me.btnVisual.TabIndex = 8
        Me.btnVisual.Text = "6 - VISUAL"
        '
        'btnCatLeak
        '
        Me.btnCatLeak.Location = New System.Drawing.Point(4, 334)
        Me.btnCatLeak.Name = "btnCatLeak"
        Me.btnCatLeak.Size = New System.Drawing.Size(94, 46)
        Me.btnCatLeak.TabIndex = 7
        Me.btnCatLeak.Text = "5.9 - CAT LEAK"
        '
        'btnPipeLeak
        '
        Me.btnPipeLeak.Location = New System.Drawing.Point(4, 287)
        Me.btnPipeLeak.Name = "btnPipeLeak"
        Me.btnPipeLeak.Size = New System.Drawing.Size(94, 46)
        Me.btnPipeLeak.TabIndex = 6
        Me.btnPipeLeak.Text = "5 - PIPE LEAK"
        '
        'btnTankLeak
        '
        Me.btnTankLeak.Location = New System.Drawing.Point(4, 240)
        Me.btnTankLeak.Name = "btnTankLeak"
        Me.btnTankLeak.Size = New System.Drawing.Size(94, 46)
        Me.btnTankLeak.TabIndex = 5
        Me.btnTankLeak.Text = "4 - TANK LEAK"
        '
        'btnCp
        '
        Me.btnCp.Location = New System.Drawing.Point(4, 194)
        Me.btnCp.Name = "btnCp"
        Me.btnCp.Size = New System.Drawing.Size(94, 46)
        Me.btnCp.TabIndex = 4
        Me.btnCp.Text = "3 - CP"
        '
        'btnSpill
        '
        Me.btnSpill.Location = New System.Drawing.Point(4, 148)
        Me.btnSpill.Name = "btnSpill"
        Me.btnSpill.Size = New System.Drawing.Size(94, 46)
        Me.btnSpill.TabIndex = 3
        Me.btnSpill.Text = "2 - SPILL"
        '
        'btnReg
        '
        Me.btnReg.Location = New System.Drawing.Point(4, 102)
        Me.btnReg.Name = "btnReg"
        Me.btnReg.Size = New System.Drawing.Size(94, 46)
        Me.btnReg.TabIndex = 2
        Me.btnReg.Text = "1 - REG"
        '
        'btnCitations
        '
        Me.btnCitations.Location = New System.Drawing.Point(4, 568)
        Me.btnCitations.Name = "btnCitations"
        Me.btnCitations.Size = New System.Drawing.Size(94, 46)
        Me.btnCitations.TabIndex = 11
        Me.btnCitations.Text = "CITATIONS"
        '
        'btnMaster
        '
        Me.btnMaster.Location = New System.Drawing.Point(4, 8)
        Me.btnMaster.Name = "btnMaster"
        Me.btnMaster.Size = New System.Drawing.Size(94, 46)
        Me.btnMaster.TabIndex = 0
        Me.btnMaster.Text = "FACILITY"
        '
        'btnTanksPipes
        '
        Me.btnTanksPipes.Location = New System.Drawing.Point(4, 55)
        Me.btnTanksPipes.Name = "btnTanksPipes"
        Me.btnTanksPipes.Size = New System.Drawing.Size(94, 46)
        Me.btnTanksPipes.TabIndex = 1
        Me.btnTanksPipes.Text = "TANKS / PIPES"
        '
        'btnMW
        '
        Me.btnMW.Location = New System.Drawing.Point(4, 520)
        Me.btnMW.Name = "btnMW"
        Me.btnMW.Size = New System.Drawing.Size(94, 46)
        Me.btnMW.TabIndex = 11
        Me.btnMW.Text = " M WELLS"
        '
        'pnlChecklistDetails
        '
        Me.pnlChecklistDetails.AutoScroll = True
        Me.pnlChecklistDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlChecklistDetails.Controls.Add(Me.pnlMW)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlMaster)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlComments)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlTOS)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlVisual)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlCatLeak)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlPipeLeak)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlTankLeak)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlCP)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlSpill)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlReg)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlInspectionCitations)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlTanksPipes)
        Me.pnlChecklistDetails.Controls.Add(Me.pnlSOC)
        Me.pnlChecklistDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlChecklistDetails.Location = New System.Drawing.Point(104, 0)
        Me.pnlChecklistDetails.Name = "pnlChecklistDetails"
        Me.pnlChecklistDetails.Size = New System.Drawing.Size(924, 669)
        Me.pnlChecklistDetails.TabIndex = 38
        '
        'pnlMW
        '
        Me.pnlMW.Controls.Add(Me.ugMW)
        Me.pnlMW.Controls.Add(Me.pnlMWAdd)
        Me.pnlMW.Controls.Add(Me.pnlMWDisplay)
        Me.pnlMW.Location = New System.Drawing.Point(16, 1672)
        Me.pnlMW.Name = "pnlMW"
        Me.pnlMW.Size = New System.Drawing.Size(856, 104)
        Me.pnlMW.TabIndex = 13
        '
        'ugMW
        '
        Me.ugMW.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugMW.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugMW.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugMW.Location = New System.Drawing.Point(0, 48)
        Me.ugMW.Name = "ugMW"
        Me.ugMW.Size = New System.Drawing.Size(856, 56)
        Me.ugMW.TabIndex = 44
        '
        'pnlMWAdd
        '
        Me.pnlMWAdd.BackColor = System.Drawing.SystemColors.Highlight
        Me.pnlMWAdd.Controls.Add(Me.btnMWAdd)
        Me.pnlMWAdd.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMWAdd.Location = New System.Drawing.Point(0, 24)
        Me.pnlMWAdd.Name = "pnlMWAdd"
        Me.pnlMWAdd.Size = New System.Drawing.Size(856, 24)
        Me.pnlMWAdd.TabIndex = 45
        Me.pnlMWAdd.Visible = False
        '
        'btnMWAdd
        '
        Me.btnMWAdd.BackColor = System.Drawing.SystemColors.Control
        Me.btnMWAdd.Enabled = False
        Me.btnMWAdd.Location = New System.Drawing.Point(104, 1)
        Me.btnMWAdd.Name = "btnMWAdd"
        Me.btnMWAdd.Size = New System.Drawing.Size(168, 23)
        Me.btnMWAdd.TabIndex = 0
        Me.btnMWAdd.Text = "Add Monitor Well Observation"
        '
        'pnlMWDisplay
        '
        Me.pnlMWDisplay.Controls.Add(Me.lblMW)
        Me.pnlMWDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMWDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlMWDisplay.Name = "pnlMWDisplay"
        Me.pnlMWDisplay.Size = New System.Drawing.Size(856, 24)
        Me.pnlMWDisplay.TabIndex = 43
        '
        'lblMW
        '
        Me.lblMW.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblMW.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblMW.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMW.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblMW.Location = New System.Drawing.Point(0, 0)
        Me.lblMW.Name = "lblMW"
        Me.lblMW.Size = New System.Drawing.Size(856, 24)
        Me.lblMW.TabIndex = 2
        Me.lblMW.Text = "MONITOR WELLS"
        Me.lblMW.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlMaster
        '
        Me.pnlMaster.Controls.Add(Me.pnlMasterDetails)
        Me.pnlMaster.Controls.Add(Me.pnlMasterDisplay)
        Me.pnlMaster.Location = New System.Drawing.Point(8, 8)
        Me.pnlMaster.Name = "pnlMaster"
        Me.pnlMaster.Size = New System.Drawing.Size(792, 220)
        Me.pnlMaster.TabIndex = 0
        Me.pnlMaster.Visible = False
        '
        'pnlMasterDetails
        '
        Me.pnlMasterDetails.AutoScroll = True
        Me.pnlMasterDetails.Controls.Add(Me.pnlContact)
        Me.pnlMasterDetails.Controls.Add(Me.pnlMasterDetailsGrid)
        Me.pnlMasterDetails.Controls.Add(Me.pnlMasterDetailsTop)
        Me.pnlMasterDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMasterDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlMasterDetails.Name = "pnlMasterDetails"
        Me.pnlMasterDetails.Size = New System.Drawing.Size(792, 196)
        Me.pnlMasterDetails.TabIndex = 3
        '
        'pnlContact
        '
        Me.pnlContact.Controls.Add(Me.Panel3)
        Me.pnlContact.Controls.Add(Me.pnlContactBottom)
        Me.pnlContact.Controls.Add(Me.pnlContactTop)
        Me.pnlContact.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlContact.Location = New System.Drawing.Point(0, 216)
        Me.pnlContact.Name = "pnlContact"
        Me.pnlContact.Size = New System.Drawing.Size(775, 232)
        Me.pnlContact.TabIndex = 61
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.ugContacts)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 32)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(775, 160)
        Me.Panel3.TabIndex = 12
        '
        'ugContacts
        '
        Me.ugContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugContacts.Name = "ugContacts"
        Me.ugContacts.Size = New System.Drawing.Size(775, 160)
        Me.ugContacts.TabIndex = 2
        '
        'pnlContactBottom
        '
        Me.pnlContactBottom.Controls.Add(Me.btnOwnerModifyContact)
        Me.pnlContactBottom.Controls.Add(Me.btnOwnerDeleteContact)
        Me.pnlContactBottom.Controls.Add(Me.btnOwnerAssociateContact)
        Me.pnlContactBottom.Controls.Add(Me.btnOwnerAddSearchContact)
        Me.pnlContactBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlContactBottom.Location = New System.Drawing.Point(0, 192)
        Me.pnlContactBottom.Name = "pnlContactBottom"
        Me.pnlContactBottom.Size = New System.Drawing.Size(775, 40)
        Me.pnlContactBottom.TabIndex = 11
        '
        'btnOwnerModifyContact
        '
        Me.btnOwnerModifyContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerModifyContact.Location = New System.Drawing.Point(328, 8)
        Me.btnOwnerModifyContact.Name = "btnOwnerModifyContact"
        Me.btnOwnerModifyContact.Size = New System.Drawing.Size(120, 23)
        Me.btnOwnerModifyContact.TabIndex = 7
        Me.btnOwnerModifyContact.Text = "Modify Contact"
        '
        'btnOwnerDeleteContact
        '
        Me.btnOwnerDeleteContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerDeleteContact.Location = New System.Drawing.Point(456, 8)
        Me.btnOwnerDeleteContact.Name = "btnOwnerDeleteContact"
        Me.btnOwnerDeleteContact.Size = New System.Drawing.Size(112, 23)
        Me.btnOwnerDeleteContact.TabIndex = 8
        Me.btnOwnerDeleteContact.Text = "Delete Contact"
        '
        'btnOwnerAssociateContact
        '
        Me.btnOwnerAssociateContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerAssociateContact.Location = New System.Drawing.Point(576, 8)
        Me.btnOwnerAssociateContact.Name = "btnOwnerAssociateContact"
        Me.btnOwnerAssociateContact.Size = New System.Drawing.Size(128, 23)
        Me.btnOwnerAssociateContact.TabIndex = 9
        Me.btnOwnerAssociateContact.Text = "Associate Contact"
        '
        'btnOwnerAddSearchContact
        '
        Me.btnOwnerAddSearchContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerAddSearchContact.Location = New System.Drawing.Point(184, 8)
        Me.btnOwnerAddSearchContact.Name = "btnOwnerAddSearchContact"
        Me.btnOwnerAddSearchContact.Size = New System.Drawing.Size(136, 23)
        Me.btnOwnerAddSearchContact.TabIndex = 6
        Me.btnOwnerAddSearchContact.Text = "Add/Search Contact"
        '
        'pnlContactTop
        '
        Me.pnlContactTop.Controls.Add(Me.lblContact)
        Me.pnlContactTop.Controls.Add(Me.chkOwnerShowContactsforAllModules)
        Me.pnlContactTop.Controls.Add(Me.chkOwnerShowRelatedContacts)
        Me.pnlContactTop.Controls.Add(Me.chkOwnerShowActiveOnly)
        Me.pnlContactTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContactTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlContactTop.Name = "pnlContactTop"
        Me.pnlContactTop.Size = New System.Drawing.Size(775, 32)
        Me.pnlContactTop.TabIndex = 10
        '
        'lblContact
        '
        Me.lblContact.Location = New System.Drawing.Point(8, 8)
        Me.lblContact.Name = "lblContact"
        Me.lblContact.Size = New System.Drawing.Size(70, 16)
        Me.lblContact.TabIndex = 6
        Me.lblContact.Text = "CONTACTS:"
        '
        'chkOwnerShowContactsforAllModules
        '
        Me.chkOwnerShowContactsforAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowContactsforAllModules.Location = New System.Drawing.Point(200, 8)
        Me.chkOwnerShowContactsforAllModules.Name = "chkOwnerShowContactsforAllModules"
        Me.chkOwnerShowContactsforAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkOwnerShowContactsforAllModules.TabIndex = 3
        Me.chkOwnerShowContactsforAllModules.Tag = "644"
        Me.chkOwnerShowContactsforAllModules.Text = "Show Contacts for All Modules"
        '
        'chkOwnerShowRelatedContacts
        '
        Me.chkOwnerShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowRelatedContacts.Location = New System.Drawing.Point(416, 8)
        Me.chkOwnerShowRelatedContacts.Name = "chkOwnerShowRelatedContacts"
        Me.chkOwnerShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkOwnerShowRelatedContacts.TabIndex = 4
        Me.chkOwnerShowRelatedContacts.Tag = "645"
        Me.chkOwnerShowRelatedContacts.Text = "Show Related Contacts"
        Me.chkOwnerShowRelatedContacts.Visible = False
        '
        'chkOwnerShowActiveOnly
        '
        Me.chkOwnerShowActiveOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowActiveOnly.Location = New System.Drawing.Point(584, 8)
        Me.chkOwnerShowActiveOnly.Name = "chkOwnerShowActiveOnly"
        Me.chkOwnerShowActiveOnly.Size = New System.Drawing.Size(144, 16)
        Me.chkOwnerShowActiveOnly.TabIndex = 5
        Me.chkOwnerShowActiveOnly.Tag = "646"
        Me.chkOwnerShowActiveOnly.Text = "Show Active Only"
        '
        'pnlMasterDetailsGrid
        '
        Me.pnlMasterDetailsGrid.Controls.Add(Me.ugInspector)
        Me.pnlMasterDetailsGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMasterDetailsGrid.Location = New System.Drawing.Point(0, 216)
        Me.pnlMasterDetailsGrid.Name = "pnlMasterDetailsGrid"
        Me.pnlMasterDetailsGrid.Size = New System.Drawing.Size(775, 232)
        Me.pnlMasterDetailsGrid.TabIndex = 60
        '
        'ugInspector
        '
        Me.ugInspector.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugInspector.Dock = System.Windows.Forms.DockStyle.Top
        Me.ugInspector.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugInspector.Location = New System.Drawing.Point(0, 0)
        Me.ugInspector.Name = "ugInspector"
        Me.ugInspector.Size = New System.Drawing.Size(775, 200)
        Me.ugInspector.TabIndex = 59
        '
        'pnlMasterDetailsTop
        '
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtOwner)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblOwner)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtFacility)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblFacility)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblOwnersPhone)
        Me.pnlMasterDetailsTop.Controls.Add(Me.mskTxtOwnerPhone)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtOwnersPhone)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtOwnerZip)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtOwnerState)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtOwnerCity)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblOwnerCity)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtOwnersAddress)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblOwnersAddress)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtOwnersRep)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblOwnersRap)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLustOwner)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtUstOwner)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLongitudeSec)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLongitudeMin)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtLongitudeSec)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtLongitudeMin)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLongitudeDegree)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtLongitudeDegree)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLogitudeW)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLatitudeSec)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtLatitudeSec)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLatitudeMin)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtLatitudeMin)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtLatitudeDegree)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLatitudeN)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtAddress)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblAddress)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLatitudeDegree)
        Me.pnlMasterDetailsTop.Controls.Add(Me.chkCapCandidate)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblState)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblZip)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblFacilityName)
        Me.pnlMasterDetailsTop.Controls.Add(Me.txtFacilityName)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLustActive)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLustActiveValue)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLustPM)
        Me.pnlMasterDetailsTop.Controls.Add(Me.lblLustPMValue)
        Me.pnlMasterDetailsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMasterDetailsTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlMasterDetailsTop.Name = "pnlMasterDetailsTop"
        Me.pnlMasterDetailsTop.Size = New System.Drawing.Size(775, 216)
        Me.pnlMasterDetailsTop.TabIndex = 0
        '
        'txtOwner
        '
        Me.txtOwner.Location = New System.Drawing.Point(120, 192)
        Me.txtOwner.Name = "txtOwner"
        Me.txtOwner.ReadOnly = True
        Me.txtOwner.TabIndex = 0
        Me.txtOwner.Text = ""
        '
        'lblOwner
        '
        Me.lblOwner.Location = New System.Drawing.Point(7, 192)
        Me.lblOwner.Name = "lblOwner"
        Me.lblOwner.Size = New System.Drawing.Size(107, 17)
        Me.lblOwner.TabIndex = 0
        Me.lblOwner.Text = "Owner Fee Balance:"
        Me.lblOwner.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtFacility
        '
        Me.txtFacility.Location = New System.Drawing.Point(472, 192)
        Me.txtFacility.Name = "txtFacility"
        Me.txtFacility.ReadOnly = True
        Me.txtFacility.TabIndex = 0
        Me.txtFacility.Text = ""
        '
        'lblFacility
        '
        Me.lblFacility.Location = New System.Drawing.Point(357, 192)
        Me.lblFacility.Name = "lblFacility"
        Me.lblFacility.Size = New System.Drawing.Size(109, 17)
        Me.lblFacility.TabIndex = 0
        Me.lblFacility.Text = "Facility Fee Balance:"
        Me.lblFacility.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblOwnersPhone
        '
        Me.lblOwnersPhone.Location = New System.Drawing.Point(30, 168)
        Me.lblOwnersPhone.Name = "lblOwnersPhone"
        Me.lblOwnersPhone.Size = New System.Drawing.Size(84, 17)
        Me.lblOwnersPhone.TabIndex = 0
        Me.lblOwnersPhone.Text = "Owner's Phone:"
        Me.lblOwnersPhone.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'mskTxtOwnerPhone
        '
        Me.mskTxtOwnerPhone.ContainingControl = Me
        Me.mskTxtOwnerPhone.Location = New System.Drawing.Point(120, 168)
        Me.mskTxtOwnerPhone.Name = "mskTxtOwnerPhone"
        Me.mskTxtOwnerPhone.OcxState = CType(resources.GetObject("mskTxtOwnerPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone.Size = New System.Drawing.Size(100, 20)
        Me.mskTxtOwnerPhone.TabIndex = 0
        '
        'txtOwnersPhone
        '
        Me.txtOwnersPhone.Location = New System.Drawing.Point(120, 168)
        Me.txtOwnersPhone.Name = "txtOwnersPhone"
        Me.txtOwnersPhone.ReadOnly = True
        Me.txtOwnersPhone.Size = New System.Drawing.Size(240, 20)
        Me.txtOwnersPhone.TabIndex = 0
        Me.txtOwnersPhone.Text = ""
        Me.txtOwnersPhone.Visible = False
        '
        'txtOwnerZip
        '
        Me.txtOwnerZip.Location = New System.Drawing.Point(568, 168)
        Me.txtOwnerZip.Name = "txtOwnerZip"
        Me.txtOwnerZip.ReadOnly = True
        Me.txtOwnerZip.Size = New System.Drawing.Size(72, 20)
        Me.txtOwnerZip.TabIndex = 0
        Me.txtOwnerZip.Text = ""
        '
        'txtOwnerState
        '
        Me.txtOwnerState.Location = New System.Drawing.Point(472, 168)
        Me.txtOwnerState.Name = "txtOwnerState"
        Me.txtOwnerState.ReadOnly = True
        Me.txtOwnerState.Size = New System.Drawing.Size(35, 20)
        Me.txtOwnerState.TabIndex = 0
        Me.txtOwnerState.Text = ""
        '
        'txtOwnerCity
        '
        Me.txtOwnerCity.Location = New System.Drawing.Point(472, 144)
        Me.txtOwnerCity.Name = "txtOwnerCity"
        Me.txtOwnerCity.ReadOnly = True
        Me.txtOwnerCity.Size = New System.Drawing.Size(168, 20)
        Me.txtOwnerCity.TabIndex = 0
        Me.txtOwnerCity.Text = ""
        '
        'lblOwnerCity
        '
        Me.lblOwnerCity.Location = New System.Drawing.Point(439, 144)
        Me.lblOwnerCity.Name = "lblOwnerCity"
        Me.lblOwnerCity.Size = New System.Drawing.Size(27, 17)
        Me.lblOwnerCity.TabIndex = 0
        Me.lblOwnerCity.Text = "City:"
        Me.lblOwnerCity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtOwnersAddress
        '
        Me.txtOwnersAddress.Location = New System.Drawing.Point(120, 144)
        Me.txtOwnersAddress.Name = "txtOwnersAddress"
        Me.txtOwnersAddress.ReadOnly = True
        Me.txtOwnersAddress.Size = New System.Drawing.Size(240, 20)
        Me.txtOwnersAddress.TabIndex = 0
        Me.txtOwnersAddress.Text = ""
        '
        'lblOwnersAddress
        '
        Me.lblOwnersAddress.Location = New System.Drawing.Point(18, 144)
        Me.lblOwnersAddress.Name = "lblOwnersAddress"
        Me.lblOwnersAddress.Size = New System.Drawing.Size(96, 17)
        Me.lblOwnersAddress.TabIndex = 0
        Me.lblOwnersAddress.Text = "Owner's Address:"
        Me.lblOwnersAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtOwnersRep
        '
        Me.txtOwnersRep.Location = New System.Drawing.Point(472, 120)
        Me.txtOwnersRep.MaxLength = 256
        Me.txtOwnersRep.Name = "txtOwnersRep"
        Me.txtOwnersRep.Size = New System.Drawing.Size(168, 20)
        Me.txtOwnersRep.TabIndex = 10
        Me.txtOwnersRep.Text = ""
        '
        'lblOwnersRap
        '
        Me.lblOwnersRap.Location = New System.Drawing.Point(394, 120)
        Me.lblOwnersRap.Name = "lblOwnersRap"
        Me.lblOwnersRap.Size = New System.Drawing.Size(72, 17)
        Me.lblOwnersRap.TabIndex = 0
        Me.lblOwnersRap.Text = "Owner's Rep:"
        Me.lblOwnersRap.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblLustOwner
        '
        Me.lblLustOwner.Location = New System.Drawing.Point(41, 120)
        Me.lblLustOwner.Name = "lblLustOwner"
        Me.lblLustOwner.Size = New System.Drawing.Size(73, 17)
        Me.lblLustOwner.TabIndex = 0
        Me.lblLustOwner.Text = "UST Owner:"
        Me.lblLustOwner.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtUstOwner
        '
        Me.txtUstOwner.Location = New System.Drawing.Point(120, 120)
        Me.txtUstOwner.Name = "txtUstOwner"
        Me.txtUstOwner.ReadOnly = True
        Me.txtUstOwner.Size = New System.Drawing.Size(240, 20)
        Me.txtUstOwner.TabIndex = 0
        Me.txtUstOwner.Text = ""
        '
        'lblLongitudeSec
        '
        Me.lblLongitudeSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.lblLongitudeSec.Location = New System.Drawing.Point(624, 56)
        Me.lblLongitudeSec.Name = "lblLongitudeSec"
        Me.lblLongitudeSec.Size = New System.Drawing.Size(13, 16)
        Me.lblLongitudeSec.TabIndex = 0
        Me.lblLongitudeSec.Text = "''"
        '
        'lblLongitudeMin
        '
        Me.lblLongitudeMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLongitudeMin.Location = New System.Drawing.Point(568, 56)
        Me.lblLongitudeMin.Name = "lblLongitudeMin"
        Me.lblLongitudeMin.Size = New System.Drawing.Size(13, 16)
        Me.lblLongitudeMin.TabIndex = 0
        Me.lblLongitudeMin.Text = "'"
        '
        'txtLongitudeSec
        '
        Me.txtLongitudeSec.Location = New System.Drawing.Point(584, 56)
        Me.txtLongitudeSec.Name = "txtLongitudeSec"
        Me.txtLongitudeSec.Size = New System.Drawing.Size(40, 20)
        Me.txtLongitudeSec.TabIndex = 9
        Me.txtLongitudeSec.Text = ""
        '
        'txtLongitudeMin
        '
        Me.txtLongitudeMin.Location = New System.Drawing.Point(528, 56)
        Me.txtLongitudeMin.Name = "txtLongitudeMin"
        Me.txtLongitudeMin.Size = New System.Drawing.Size(40, 20)
        Me.txtLongitudeMin.TabIndex = 8
        Me.txtLongitudeMin.Text = ""
        '
        'lblLongitudeDegree
        '
        Me.lblLongitudeDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.lblLongitudeDegree.Location = New System.Drawing.Point(512, 48)
        Me.lblLongitudeDegree.Name = "lblLongitudeDegree"
        Me.lblLongitudeDegree.Size = New System.Drawing.Size(13, 16)
        Me.lblLongitudeDegree.TabIndex = 0
        Me.lblLongitudeDegree.Text = "o"
        '
        'txtLongitudeDegree
        '
        Me.txtLongitudeDegree.Location = New System.Drawing.Point(472, 56)
        Me.txtLongitudeDegree.Name = "txtLongitudeDegree"
        Me.txtLongitudeDegree.Size = New System.Drawing.Size(40, 20)
        Me.txtLongitudeDegree.TabIndex = 7
        Me.txtLongitudeDegree.Text = ""
        '
        'lblLogitudeW
        '
        Me.lblLogitudeW.Location = New System.Drawing.Point(395, 56)
        Me.lblLogitudeW.Name = "lblLogitudeW"
        Me.lblLogitudeW.Size = New System.Drawing.Size(71, 17)
        Me.lblLogitudeW.TabIndex = 0
        Me.lblLogitudeW.Text = "Longitude W:"
        Me.lblLogitudeW.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblLatitudeSec
        '
        Me.lblLatitudeSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLatitudeSec.Location = New System.Drawing.Point(624, 32)
        Me.lblLatitudeSec.Name = "lblLatitudeSec"
        Me.lblLatitudeSec.Size = New System.Drawing.Size(13, 16)
        Me.lblLatitudeSec.TabIndex = 0
        Me.lblLatitudeSec.Text = "''"
        '
        'txtLatitudeSec
        '
        Me.txtLatitudeSec.Location = New System.Drawing.Point(584, 32)
        Me.txtLatitudeSec.Name = "txtLatitudeSec"
        Me.txtLatitudeSec.Size = New System.Drawing.Size(40, 20)
        Me.txtLatitudeSec.TabIndex = 6
        Me.txtLatitudeSec.Text = ""
        '
        'lblLatitudeMin
        '
        Me.lblLatitudeMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLatitudeMin.Location = New System.Drawing.Point(568, 32)
        Me.lblLatitudeMin.Name = "lblLatitudeMin"
        Me.lblLatitudeMin.Size = New System.Drawing.Size(13, 16)
        Me.lblLatitudeMin.TabIndex = 0
        Me.lblLatitudeMin.Text = "'"
        '
        'txtLatitudeMin
        '
        Me.txtLatitudeMin.Location = New System.Drawing.Point(528, 32)
        Me.txtLatitudeMin.Name = "txtLatitudeMin"
        Me.txtLatitudeMin.Size = New System.Drawing.Size(40, 20)
        Me.txtLatitudeMin.TabIndex = 5
        Me.txtLatitudeMin.Text = ""
        '
        'txtLatitudeDegree
        '
        Me.txtLatitudeDegree.Location = New System.Drawing.Point(472, 32)
        Me.txtLatitudeDegree.Name = "txtLatitudeDegree"
        Me.txtLatitudeDegree.Size = New System.Drawing.Size(40, 20)
        Me.txtLatitudeDegree.TabIndex = 4
        Me.txtLatitudeDegree.Text = ""
        '
        'lblLatitudeN
        '
        Me.lblLatitudeN.Location = New System.Drawing.Point(408, 32)
        Me.lblLatitudeN.Name = "lblLatitudeN"
        Me.lblLatitudeN.Size = New System.Drawing.Size(59, 17)
        Me.lblLatitudeN.TabIndex = 0
        Me.lblLatitudeN.Text = "Latitude N:"
        Me.lblLatitudeN.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(120, 32)
        Me.txtAddress.Multiline = True
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.ReadOnly = True
        Me.txtAddress.Size = New System.Drawing.Size(240, 83)
        Me.txtAddress.TabIndex = 0
        Me.txtAddress.Text = ""
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(28, 32)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(87, 17)
        Me.lblAddress.TabIndex = 0
        Me.lblAddress.Text = "Facility Address:"
        Me.lblAddress.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblLatitudeDegree
        '
        Me.lblLatitudeDegree.BackColor = System.Drawing.SystemColors.Control
        Me.lblLatitudeDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.lblLatitudeDegree.Location = New System.Drawing.Point(512, 24)
        Me.lblLatitudeDegree.Name = "lblLatitudeDegree"
        Me.lblLatitudeDegree.Size = New System.Drawing.Size(13, 16)
        Me.lblLatitudeDegree.TabIndex = 0
        Me.lblLatitudeDegree.Text = "o"
        '
        'chkCapCandidate
        '
        Me.chkCapCandidate.Location = New System.Drawing.Point(371, 88)
        Me.chkCapCandidate.Name = "chkCapCandidate"
        Me.chkCapCandidate.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkCapCandidate.Size = New System.Drawing.Size(115, 24)
        Me.chkCapCandidate.TabIndex = 0
        Me.chkCapCandidate.Text = " :CAP Candidate"
        '
        'lblState
        '
        Me.lblState.Location = New System.Drawing.Point(432, 168)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(34, 17)
        Me.lblState.TabIndex = 0
        Me.lblState.Text = "State:"
        Me.lblState.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblZip
        '
        Me.lblZip.Location = New System.Drawing.Point(537, 168)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(25, 17)
        Me.lblZip.TabIndex = 0
        Me.lblZip.Text = "Zip:"
        Me.lblZip.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblFacilityName
        '
        Me.lblFacilityName.Location = New System.Drawing.Point(32, 8)
        Me.lblFacilityName.Name = "lblFacilityName"
        Me.lblFacilityName.Size = New System.Drawing.Size(80, 17)
        Me.lblFacilityName.TabIndex = 0
        Me.lblFacilityName.Text = "Facility name:"
        Me.lblFacilityName.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtFacilityName
        '
        Me.txtFacilityName.Location = New System.Drawing.Point(120, 8)
        Me.txtFacilityName.MaxLength = 256
        Me.txtFacilityName.Name = "txtFacilityName"
        Me.txtFacilityName.Size = New System.Drawing.Size(240, 20)
        Me.txtFacilityName.TabIndex = 10
        Me.txtFacilityName.Text = ""
        '
        'lblLustActive
        '
        Me.lblLustActive.Location = New System.Drawing.Point(400, 8)
        Me.lblLustActive.Name = "lblLustActive"
        Me.lblLustActive.Size = New System.Drawing.Size(64, 17)
        Me.lblLustActive.TabIndex = 0
        Me.lblLustActive.Text = "Lust Active:"
        Me.lblLustActive.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblLustActiveValue
        '
        Me.lblLustActiveValue.Location = New System.Drawing.Point(472, 8)
        Me.lblLustActiveValue.Name = "lblLustActiveValue"
        Me.lblLustActiveValue.Size = New System.Drawing.Size(32, 17)
        Me.lblLustActiveValue.TabIndex = 0
        '
        'lblLustPM
        '
        Me.lblLustPM.Location = New System.Drawing.Point(504, 8)
        Me.lblLustPM.Name = "lblLustPM"
        Me.lblLustPM.Size = New System.Drawing.Size(56, 17)
        Me.lblLustPM.TabIndex = 0
        Me.lblLustPM.Text = "Lust PM:"
        Me.lblLustPM.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblLustPMValue
        '
        Me.lblLustPMValue.Location = New System.Drawing.Point(568, 8)
        Me.lblLustPMValue.Name = "lblLustPMValue"
        Me.lblLustPMValue.Size = New System.Drawing.Size(184, 17)
        Me.lblLustPMValue.TabIndex = 0
        '
        'pnlMasterDisplay
        '
        Me.pnlMasterDisplay.Controls.Add(Me.lblMasterDisplay)
        Me.pnlMasterDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMasterDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlMasterDisplay.Name = "pnlMasterDisplay"
        Me.pnlMasterDisplay.Size = New System.Drawing.Size(792, 24)
        Me.pnlMasterDisplay.TabIndex = 4
        '
        'lblMasterDisplay
        '
        Me.lblMasterDisplay.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblMasterDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblMasterDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMasterDisplay.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblMasterDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblMasterDisplay.Name = "lblMasterDisplay"
        Me.lblMasterDisplay.Size = New System.Drawing.Size(792, 24)
        Me.lblMasterDisplay.TabIndex = 2
        Me.lblMasterDisplay.Text = "Master"
        Me.lblMasterDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlComments
        '
        Me.pnlComments.Controls.Add(Me.txtComments)
        Me.pnlComments.Controls.Add(Me.pnlCommentsDisplay)
        Me.pnlComments.Location = New System.Drawing.Point(16, 1552)
        Me.pnlComments.Name = "pnlComments"
        Me.pnlComments.Size = New System.Drawing.Size(864, 112)
        Me.pnlComments.TabIndex = 11
        '
        'txtComments
        '
        Me.txtComments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtComments.Location = New System.Drawing.Point(0, 24)
        Me.txtComments.MaxLength = 2147483647
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtComments.Size = New System.Drawing.Size(864, 88)
        Me.txtComments.TabIndex = 44
        Me.txtComments.Text = ""
        '
        'pnlCommentsDisplay
        '
        Me.pnlCommentsDisplay.Controls.Add(Me.lblComments)
        Me.pnlCommentsDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCommentsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlCommentsDisplay.Name = "pnlCommentsDisplay"
        Me.pnlCommentsDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlCommentsDisplay.TabIndex = 43
        '
        'lblComments
        '
        Me.lblComments.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblComments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblComments.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblComments.Location = New System.Drawing.Point(0, 0)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(864, 24)
        Me.lblComments.TabIndex = 2
        Me.lblComments.Text = "8 - INSPECTION COMMENTS"
        Me.lblComments.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlTOS
        '
        Me.pnlTOS.Controls.Add(Me.ugTOS)
        Me.pnlTOS.Controls.Add(Me.pnlTOSDisplay)
        Me.pnlTOS.Location = New System.Drawing.Point(24, 1440)
        Me.pnlTOS.Name = "pnlTOS"
        Me.pnlTOS.Size = New System.Drawing.Size(856, 104)
        Me.pnlTOS.TabIndex = 10
        '
        'ugTOS
        '
        Me.ugTOS.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTOS.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugTOS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugTOS.Location = New System.Drawing.Point(0, 24)
        Me.ugTOS.Name = "ugTOS"
        Me.ugTOS.Size = New System.Drawing.Size(856, 80)
        Me.ugTOS.TabIndex = 44
        '
        'pnlTOSDisplay
        '
        Me.pnlTOSDisplay.Controls.Add(Me.lblTOS)
        Me.pnlTOSDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTOSDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlTOSDisplay.Name = "pnlTOSDisplay"
        Me.pnlTOSDisplay.Size = New System.Drawing.Size(856, 24)
        Me.pnlTOSDisplay.TabIndex = 43
        '
        'lblTOS
        '
        Me.lblTOS.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTOS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTOS.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTOS.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTOS.Location = New System.Drawing.Point(0, 0)
        Me.lblTOS.Name = "lblTOS"
        Me.lblTOS.Size = New System.Drawing.Size(856, 24)
        Me.lblTOS.TabIndex = 2
        Me.lblTOS.Text = "7 - TEMPORARILY OUT OF SERVICE TANKS"
        Me.lblTOS.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlVisual
        '
        Me.pnlVisual.Controls.Add(Me.ugVisual)
        Me.pnlVisual.Controls.Add(Me.pnlVisualDisplay)
        Me.pnlVisual.Location = New System.Drawing.Point(16, 1328)
        Me.pnlVisual.Name = "pnlVisual"
        Me.pnlVisual.Size = New System.Drawing.Size(864, 104)
        Me.pnlVisual.TabIndex = 9
        '
        'ugVisual
        '
        Me.ugVisual.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugVisual.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugVisual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugVisual.Location = New System.Drawing.Point(0, 24)
        Me.ugVisual.Name = "ugVisual"
        Me.ugVisual.Size = New System.Drawing.Size(864, 80)
        Me.ugVisual.TabIndex = 43
        '
        'pnlVisualDisplay
        '
        Me.pnlVisualDisplay.Controls.Add(Me.lblVisual)
        Me.pnlVisualDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlVisualDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlVisualDisplay.Name = "pnlVisualDisplay"
        Me.pnlVisualDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlVisualDisplay.TabIndex = 42
        '
        'lblVisual
        '
        Me.lblVisual.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblVisual.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblVisual.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVisual.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblVisual.Location = New System.Drawing.Point(0, 0)
        Me.lblVisual.Name = "lblVisual"
        Me.lblVisual.Size = New System.Drawing.Size(864, 24)
        Me.lblVisual.TabIndex = 2
        Me.lblVisual.Text = "6 - INSPECTOR'S VISUAL OBSERVATIONS"
        Me.lblVisual.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlCatLeak
        '
        Me.pnlCatLeak.Controls.Add(Me.ugCatLeak)
        Me.pnlCatLeak.Controls.Add(Me.pnlCatLeakDisplay)
        Me.pnlCatLeak.Location = New System.Drawing.Point(8, 1184)
        Me.pnlCatLeak.Name = "pnlCatLeak"
        Me.pnlCatLeak.Size = New System.Drawing.Size(864, 136)
        Me.pnlCatLeak.TabIndex = 8
        '
        'ugCatLeak
        '
        Me.ugCatLeak.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCatLeak.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCatLeak.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCatLeak.Location = New System.Drawing.Point(0, 24)
        Me.ugCatLeak.Name = "ugCatLeak"
        Me.ugCatLeak.Size = New System.Drawing.Size(864, 112)
        Me.ugCatLeak.TabIndex = 42
        '
        'pnlCatLeakDisplay
        '
        Me.pnlCatLeakDisplay.Controls.Add(Me.lblCatLeak)
        Me.pnlCatLeakDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCatLeakDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlCatLeakDisplay.Name = "pnlCatLeakDisplay"
        Me.pnlCatLeakDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlCatLeakDisplay.TabIndex = 41
        '
        'lblCatLeak
        '
        Me.lblCatLeak.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblCatLeak.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCatLeak.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCatLeak.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCatLeak.Location = New System.Drawing.Point(0, 0)
        Me.lblCatLeak.Name = "lblCatLeak"
        Me.lblCatLeak.Size = New System.Drawing.Size(864, 24)
        Me.lblCatLeak.TabIndex = 2
        Me.lblCatLeak.Text = "5.9 - PRESSURIZED PIPING CATASTROPHIC LEAK DETECTION"
        Me.lblCatLeak.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlPipeLeak
        '
        Me.pnlPipeLeak.Controls.Add(Me.ugPipeLeak)
        Me.pnlPipeLeak.Controls.Add(Me.pnlPipeLeakAdd)
        Me.pnlPipeLeak.Controls.Add(Me.pnlPipeLeakDisplay)
        Me.pnlPipeLeak.Location = New System.Drawing.Point(8, 1024)
        Me.pnlPipeLeak.Name = "pnlPipeLeak"
        Me.pnlPipeLeak.Size = New System.Drawing.Size(864, 152)
        Me.pnlPipeLeak.TabIndex = 7
        '
        'ugPipeLeak
        '
        Me.ugPipeLeak.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPipeLeak.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugPipeLeak.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugPipeLeak.Location = New System.Drawing.Point(0, 48)
        Me.ugPipeLeak.Name = "ugPipeLeak"
        Me.ugPipeLeak.Size = New System.Drawing.Size(864, 104)
        Me.ugPipeLeak.TabIndex = 41
        '
        'pnlPipeLeakAdd
        '
        Me.pnlPipeLeakAdd.BackColor = System.Drawing.SystemColors.Highlight
        Me.pnlPipeLeakAdd.Controls.Add(Me.btnAddPipeMW)
        Me.pnlPipeLeakAdd.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeLeakAdd.Location = New System.Drawing.Point(0, 24)
        Me.pnlPipeLeakAdd.Name = "pnlPipeLeakAdd"
        Me.pnlPipeLeakAdd.Size = New System.Drawing.Size(864, 24)
        Me.pnlPipeLeakAdd.TabIndex = 45
        Me.pnlPipeLeakAdd.Visible = False
        '
        'btnAddPipeMW
        '
        Me.btnAddPipeMW.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddPipeMW.Enabled = False
        Me.btnAddPipeMW.Location = New System.Drawing.Point(104, 1)
        Me.btnAddPipeMW.Name = "btnAddPipeMW"
        Me.btnAddPipeMW.Size = New System.Drawing.Size(168, 23)
        Me.btnAddPipeMW.TabIndex = 0
        Me.btnAddPipeMW.Text = "Add Monitor Well Observation"
        '
        'pnlPipeLeakDisplay
        '
        Me.pnlPipeLeakDisplay.Controls.Add(Me.lblPipeLeak)
        Me.pnlPipeLeakDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeLeakDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlPipeLeakDisplay.Name = "pnlPipeLeakDisplay"
        Me.pnlPipeLeakDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlPipeLeakDisplay.TabIndex = 40
        '
        'lblPipeLeak
        '
        Me.lblPipeLeak.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPipeLeak.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPipeLeak.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPipeLeak.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPipeLeak.Location = New System.Drawing.Point(0, 0)
        Me.lblPipeLeak.Name = "lblPipeLeak"
        Me.lblPipeLeak.Size = New System.Drawing.Size(864, 24)
        Me.lblPipeLeak.TabIndex = 2
        Me.lblPipeLeak.Text = "5 - PIPING LEAK DETECTION - PRIMARY"
        Me.lblPipeLeak.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlTankLeak
        '
        Me.pnlTankLeak.Controls.Add(Me.ugTankLeak)
        Me.pnlTankLeak.Controls.Add(Me.pnlTankLeakAdd)
        Me.pnlTankLeak.Controls.Add(Me.pnlTankLeakDisplay)
        Me.pnlTankLeak.Location = New System.Drawing.Point(8, 872)
        Me.pnlTankLeak.Name = "pnlTankLeak"
        Me.pnlTankLeak.Size = New System.Drawing.Size(864, 144)
        Me.pnlTankLeak.TabIndex = 6
        '
        'ugTankLeak
        '
        Me.ugTankLeak.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTankLeak.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugTankLeak.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugTankLeak.Location = New System.Drawing.Point(0, 48)
        Me.ugTankLeak.Name = "ugTankLeak"
        Me.ugTankLeak.Size = New System.Drawing.Size(864, 96)
        Me.ugTankLeak.TabIndex = 40
        '
        'pnlTankLeakAdd
        '
        Me.pnlTankLeakAdd.BackColor = System.Drawing.SystemColors.Highlight
        Me.pnlTankLeakAdd.Controls.Add(Me.btnAddTankMW)
        Me.pnlTankLeakAdd.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankLeakAdd.Location = New System.Drawing.Point(0, 24)
        Me.pnlTankLeakAdd.Name = "pnlTankLeakAdd"
        Me.pnlTankLeakAdd.Size = New System.Drawing.Size(864, 24)
        Me.pnlTankLeakAdd.TabIndex = 44
        Me.pnlTankLeakAdd.Visible = False
        '
        'btnAddTankMW
        '
        Me.btnAddTankMW.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddTankMW.Enabled = False
        Me.btnAddTankMW.Location = New System.Drawing.Point(104, 1)
        Me.btnAddTankMW.Name = "btnAddTankMW"
        Me.btnAddTankMW.Size = New System.Drawing.Size(168, 23)
        Me.btnAddTankMW.TabIndex = 0
        Me.btnAddTankMW.Text = "Add Monitor Well Observation"
        '
        'pnlTankLeakDisplay
        '
        Me.pnlTankLeakDisplay.Controls.Add(Me.lblTankLeak)
        Me.pnlTankLeakDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTankLeakDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlTankLeakDisplay.Name = "pnlTankLeakDisplay"
        Me.pnlTankLeakDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlTankLeakDisplay.TabIndex = 39
        '
        'lblTankLeak
        '
        Me.lblTankLeak.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTankLeak.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTankLeak.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTankLeak.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTankLeak.Location = New System.Drawing.Point(0, 0)
        Me.lblTankLeak.Name = "lblTankLeak"
        Me.lblTankLeak.Size = New System.Drawing.Size(864, 24)
        Me.lblTankLeak.TabIndex = 2
        Me.lblTankLeak.Text = "4 - TANK LEAK DETECTION"
        Me.lblTankLeak.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlCP
        '
        Me.pnlCP.Controls.Add(Me.ugCP)
        Me.pnlCP.Controls.Add(Me.pnlCPAdd)
        Me.pnlCP.Controls.Add(Me.pnlCPDisplay)
        Me.pnlCP.Location = New System.Drawing.Point(8, 688)
        Me.pnlCP.Name = "pnlCP"
        Me.pnlCP.Size = New System.Drawing.Size(864, 176)
        Me.pnlCP.TabIndex = 5
        '
        'ugCP
        '
        Me.ugCP.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCP.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCP.Location = New System.Drawing.Point(0, 48)
        Me.ugCP.Name = "ugCP"
        Me.ugCP.Size = New System.Drawing.Size(864, 128)
        Me.ugCP.TabIndex = 42
        '
        'pnlCPAdd
        '
        Me.pnlCPAdd.BackColor = System.Drawing.SystemColors.Highlight
        Me.pnlCPAdd.Controls.Add(Me.btnAddTankCP)
        Me.pnlCPAdd.Controls.Add(Me.btnAddPipeCP)
        Me.pnlCPAdd.Controls.Add(Me.btnAddTermCP)
        Me.pnlCPAdd.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCPAdd.Location = New System.Drawing.Point(0, 24)
        Me.pnlCPAdd.Name = "pnlCPAdd"
        Me.pnlCPAdd.Size = New System.Drawing.Size(864, 24)
        Me.pnlCPAdd.TabIndex = 43
        Me.pnlCPAdd.Visible = False
        '
        'btnAddTankCP
        '
        Me.btnAddTankCP.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddTankCP.Enabled = False
        Me.btnAddTankCP.Location = New System.Drawing.Point(104, 1)
        Me.btnAddTankCP.Name = "btnAddTankCP"
        Me.btnAddTankCP.Size = New System.Drawing.Size(128, 23)
        Me.btnAddTankCP.TabIndex = 0
        Me.btnAddTankCP.Text = "Add Tank CP Reading"
        '
        'btnAddPipeCP
        '
        Me.btnAddPipeCP.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddPipeCP.Enabled = False
        Me.btnAddPipeCP.Location = New System.Drawing.Point(234, 1)
        Me.btnAddPipeCP.Name = "btnAddPipeCP"
        Me.btnAddPipeCP.Size = New System.Drawing.Size(128, 23)
        Me.btnAddPipeCP.TabIndex = 0
        Me.btnAddPipeCP.Text = "Add Pipe CP Reading"
        '
        'btnAddTermCP
        '
        Me.btnAddTermCP.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddTermCP.Enabled = False
        Me.btnAddTermCP.Location = New System.Drawing.Point(364, 1)
        Me.btnAddTermCP.Name = "btnAddTermCP"
        Me.btnAddTermCP.Size = New System.Drawing.Size(128, 23)
        Me.btnAddTermCP.TabIndex = 0
        Me.btnAddTermCP.Text = "Add Term CP Reading"
        '
        'pnlCPDisplay
        '
        Me.pnlCPDisplay.Controls.Add(Me.lblCP)
        Me.pnlCPDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCPDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlCPDisplay.Name = "pnlCPDisplay"
        Me.pnlCPDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlCPDisplay.TabIndex = 39
        '
        'lblCP
        '
        Me.lblCP.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblCP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCP.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCP.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCP.Location = New System.Drawing.Point(0, 0)
        Me.lblCP.Name = "lblCP"
        Me.lblCP.Size = New System.Drawing.Size(864, 24)
        Me.lblCP.TabIndex = 2
        Me.lblCP.Text = "3 - CORROSION PROTECTION"
        Me.lblCP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlSpill
        '
        Me.pnlSpill.Controls.Add(Me.ugSpill)
        Me.pnlSpill.Controls.Add(Me.pnlSpillDisplay)
        Me.pnlSpill.Location = New System.Drawing.Point(8, 528)
        Me.pnlSpill.Name = "pnlSpill"
        Me.pnlSpill.Size = New System.Drawing.Size(864, 144)
        Me.pnlSpill.TabIndex = 4
        '
        'ugSpill
        '
        Me.ugSpill.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugSpill.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugSpill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugSpill.Location = New System.Drawing.Point(0, 24)
        Me.ugSpill.Name = "ugSpill"
        Me.ugSpill.Size = New System.Drawing.Size(864, 120)
        Me.ugSpill.TabIndex = 38
        '
        'pnlSpillDisplay
        '
        Me.pnlSpillDisplay.Controls.Add(Me.lblSpill)
        Me.pnlSpillDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlSpillDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlSpillDisplay.Name = "pnlSpillDisplay"
        Me.pnlSpillDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlSpillDisplay.TabIndex = 37
        '
        'lblSpill
        '
        Me.lblSpill.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblSpill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblSpill.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSpill.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblSpill.Location = New System.Drawing.Point(0, 0)
        Me.lblSpill.Name = "lblSpill"
        Me.lblSpill.Size = New System.Drawing.Size(864, 24)
        Me.lblSpill.TabIndex = 2
        Me.lblSpill.Text = "2 - SPILL AND OVERFILL PREVENTION"
        Me.lblSpill.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlReg
        '
        Me.pnlReg.Controls.Add(Me.pnlRegDisplay)
        Me.pnlReg.Controls.Add(Me.ugReg)
        Me.pnlReg.Location = New System.Drawing.Point(8, 424)
        Me.pnlReg.Name = "pnlReg"
        Me.pnlReg.Size = New System.Drawing.Size(864, 96)
        Me.pnlReg.TabIndex = 3
        Me.pnlReg.Visible = False
        '
        'pnlRegDisplay
        '
        Me.pnlRegDisplay.Controls.Add(Me.lblReg)
        Me.pnlRegDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlRegDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlRegDisplay.Name = "pnlRegDisplay"
        Me.pnlRegDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlRegDisplay.TabIndex = 36
        '
        'lblReg
        '
        Me.lblReg.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblReg.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblReg.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReg.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblReg.Location = New System.Drawing.Point(0, 0)
        Me.lblReg.Name = "lblReg"
        Me.lblReg.Size = New System.Drawing.Size(864, 24)
        Me.lblReg.TabIndex = 2
        Me.lblReg.Text = "1 - REGISTRATION/TESTING/CONSTRUCTION"
        Me.lblReg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ugReg
        '
        Me.ugReg.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugReg.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugReg.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugReg.Location = New System.Drawing.Point(0, 0)
        Me.ugReg.Name = "ugReg"
        Me.ugReg.Size = New System.Drawing.Size(864, 96)
        Me.ugReg.TabIndex = 36
        Me.ugReg.Text = "Question"
        '
        'pnlInspectionCitations
        '
        Me.pnlInspectionCitations.Controls.Add(Me.pnlInspecCitationsDisplay)
        Me.pnlInspectionCitations.Controls.Add(Me.ugCitation)
        Me.pnlInspectionCitations.Location = New System.Drawing.Point(8, 304)
        Me.pnlInspectionCitations.Name = "pnlInspectionCitations"
        Me.pnlInspectionCitations.Size = New System.Drawing.Size(864, 112)
        Me.pnlInspectionCitations.TabIndex = 2
        Me.pnlInspectionCitations.Visible = False
        '
        'pnlInspecCitationsDisplay
        '
        Me.pnlInspecCitationsDisplay.Controls.Add(Me.lblInspecCitationsDisplay)
        Me.pnlInspecCitationsDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlInspecCitationsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlInspecCitationsDisplay.Name = "pnlInspecCitationsDisplay"
        Me.pnlInspecCitationsDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlInspecCitationsDisplay.TabIndex = 34
        '
        'lblInspecCitationsDisplay
        '
        Me.lblInspecCitationsDisplay.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblInspecCitationsDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblInspecCitationsDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInspecCitationsDisplay.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblInspecCitationsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblInspecCitationsDisplay.Name = "lblInspecCitationsDisplay"
        Me.lblInspecCitationsDisplay.Size = New System.Drawing.Size(864, 24)
        Me.lblInspecCitationsDisplay.TabIndex = 2
        Me.lblInspecCitationsDisplay.Text = "INSPECTION CITATIONS"
        Me.lblInspecCitationsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ugCitation
        '
        Me.ugCitation.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCitation.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCitation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCitation.Location = New System.Drawing.Point(0, 0)
        Me.ugCitation.Name = "ugCitation"
        Me.ugCitation.Size = New System.Drawing.Size(864, 112)
        Me.ugCitation.TabIndex = 34
        Me.ugCitation.Text = "Citation"
        '
        'pnlTanksPipes
        '
        Me.pnlTanksPipes.Controls.Add(Me.pnlTanksPipesDetails)
        Me.pnlTanksPipes.Controls.Add(Me.pnlTanksPipesDisplay)
        Me.pnlTanksPipes.Location = New System.Drawing.Point(8, 8)
        Me.pnlTanksPipes.Name = "pnlTanksPipes"
        Me.pnlTanksPipes.Size = New System.Drawing.Size(864, 288)
        Me.pnlTanksPipes.TabIndex = 1
        Me.pnlTanksPipes.Visible = False
        '
        'pnlTanksPipesDetails
        '
        Me.pnlTanksPipesDetails.AutoScroll = True
        Me.pnlTanksPipesDetails.Controls.Add(Me.pnlTerminationsDetails)
        Me.pnlTanksPipesDetails.Controls.Add(Me.pnlTerminationsTop)
        Me.pnlTanksPipesDetails.Controls.Add(Me.pnlPipeDetails)
        Me.pnlTanksPipesDetails.Controls.Add(Me.pnlPipeTop)
        Me.pnlTanksPipesDetails.Controls.Add(Me.pnlTanksDetails)
        Me.pnlTanksPipesDetails.Controls.Add(Me.pnlTanksTop)
        Me.pnlTanksPipesDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlTanksPipesDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlTanksPipesDetails.Name = "pnlTanksPipesDetails"
        Me.pnlTanksPipesDetails.Size = New System.Drawing.Size(864, 264)
        Me.pnlTanksPipesDetails.TabIndex = 31
        '
        'pnlTerminationsDetails
        '
        Me.pnlTerminationsDetails.Controls.Add(Me.ugTerminations)
        Me.pnlTerminationsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTerminationsDetails.Location = New System.Drawing.Point(0, 407)
        Me.pnlTerminationsDetails.Name = "pnlTerminationsDetails"
        Me.pnlTerminationsDetails.Size = New System.Drawing.Size(847, 162)
        Me.pnlTerminationsDetails.TabIndex = 38
        '
        'ugTerminations
        '
        Me.ugTerminations.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTerminations.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugTerminations.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugTerminations.Location = New System.Drawing.Point(0, 0)
        Me.ugTerminations.Name = "ugTerminations"
        Me.ugTerminations.Size = New System.Drawing.Size(847, 162)
        Me.ugTerminations.TabIndex = 32
        '
        'pnlTerminationsTop
        '
        Me.pnlTerminationsTop.Controls.Add(Me.lblTerminations)
        Me.pnlTerminationsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTerminationsTop.Location = New System.Drawing.Point(0, 384)
        Me.pnlTerminationsTop.Name = "pnlTerminationsTop"
        Me.pnlTerminationsTop.Size = New System.Drawing.Size(847, 23)
        Me.pnlTerminationsTop.TabIndex = 37
        '
        'lblTerminations
        '
        Me.lblTerminations.Location = New System.Drawing.Point(16, 5)
        Me.lblTerminations.Name = "lblTerminations"
        Me.lblTerminations.Size = New System.Drawing.Size(72, 17)
        Me.lblTerminations.TabIndex = 4
        Me.lblTerminations.Text = "Terminations"
        '
        'pnlPipeDetails
        '
        Me.pnlPipeDetails.Controls.Add(Me.ugPipes)
        Me.pnlPipeDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeDetails.Location = New System.Drawing.Point(0, 215)
        Me.pnlPipeDetails.Name = "pnlPipeDetails"
        Me.pnlPipeDetails.Size = New System.Drawing.Size(847, 169)
        Me.pnlPipeDetails.TabIndex = 36
        '
        'ugPipes
        '
        Me.ugPipes.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPipes.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugPipes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugPipes.Location = New System.Drawing.Point(0, 0)
        Me.ugPipes.Name = "ugPipes"
        Me.ugPipes.Size = New System.Drawing.Size(847, 169)
        Me.ugPipes.TabIndex = 31
        '
        'pnlPipeTop
        '
        Me.pnlPipeTop.Controls.Add(Me.lblPipes)
        Me.pnlPipeTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipeTop.Location = New System.Drawing.Point(0, 192)
        Me.pnlPipeTop.Name = "pnlPipeTop"
        Me.pnlPipeTop.Size = New System.Drawing.Size(847, 23)
        Me.pnlPipeTop.TabIndex = 35
        '
        'lblPipes
        '
        Me.lblPipes.Location = New System.Drawing.Point(8, 3)
        Me.lblPipes.Name = "lblPipes"
        Me.lblPipes.Size = New System.Drawing.Size(40, 17)
        Me.lblPipes.TabIndex = 2
        Me.lblPipes.Text = "Pipes"
        '
        'pnlTanksDetails
        '
        Me.pnlTanksDetails.Controls.Add(Me.ugTanks)
        Me.pnlTanksDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTanksDetails.Location = New System.Drawing.Point(0, 23)
        Me.pnlTanksDetails.Name = "pnlTanksDetails"
        Me.pnlTanksDetails.Size = New System.Drawing.Size(847, 169)
        Me.pnlTanksDetails.TabIndex = 34
        '
        'ugTanks
        '
        Me.ugTanks.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTanks.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugTanks.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugTanks.Location = New System.Drawing.Point(0, 0)
        Me.ugTanks.Name = "ugTanks"
        Me.ugTanks.Size = New System.Drawing.Size(847, 169)
        Me.ugTanks.TabIndex = 30
        '
        'pnlTanksTop
        '
        Me.pnlTanksTop.Controls.Add(Me.lblTanks)
        Me.pnlTanksTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTanksTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTanksTop.Name = "pnlTanksTop"
        Me.pnlTanksTop.Size = New System.Drawing.Size(847, 23)
        Me.pnlTanksTop.TabIndex = 33
        '
        'lblTanks
        '
        Me.lblTanks.Location = New System.Drawing.Point(6, 5)
        Me.lblTanks.Name = "lblTanks"
        Me.lblTanks.Size = New System.Drawing.Size(40, 17)
        Me.lblTanks.TabIndex = 0
        Me.lblTanks.Text = "Tanks"
        '
        'pnlTanksPipesDisplay
        '
        Me.pnlTanksPipesDisplay.Controls.Add(Me.lblTanksPipesDisplay)
        Me.pnlTanksPipesDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTanksPipesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlTanksPipesDisplay.Name = "pnlTanksPipesDisplay"
        Me.pnlTanksPipesDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlTanksPipesDisplay.TabIndex = 30
        '
        'lblTanksPipesDisplay
        '
        Me.lblTanksPipesDisplay.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTanksPipesDisplay.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTanksPipesDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTanksPipesDisplay.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTanksPipesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblTanksPipesDisplay.Name = "lblTanksPipesDisplay"
        Me.lblTanksPipesDisplay.Size = New System.Drawing.Size(864, 24)
        Me.lblTanksPipesDisplay.TabIndex = 2
        Me.lblTanksPipesDisplay.Text = "TANKS/PIPES"
        Me.lblTanksPipesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlSOC
        '
        Me.pnlSOC.Controls.Add(Me.Panel2)
        Me.pnlSOC.Controls.Add(Me.Panel1)
        Me.pnlSOC.Controls.Add(Me.pnlSOCDisplay)
        Me.pnlSOC.Location = New System.Drawing.Point(16, 1784)
        Me.pnlSOC.Name = "pnlSOC"
        Me.pnlSOC.Size = New System.Drawing.Size(864, 104)
        Me.pnlSOC.TabIndex = 12
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.ugSOC)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 56)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(864, 48)
        Me.Panel2.TabIndex = 48
        '
        'ugSOC
        '
        Me.ugSOC.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugSOC.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugSOC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugSOC.Location = New System.Drawing.Point(0, 0)
        Me.ugSOC.Name = "ugSOC"
        Me.ugSOC.Size = New System.Drawing.Size(864, 48)
        Me.ugSOC.TabIndex = 45
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.chkSOCSketchAttached)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 24)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(864, 32)
        Me.Panel1.TabIndex = 46
        '
        'chkSOCSketchAttached
        '
        Me.chkSOCSketchAttached.Location = New System.Drawing.Point(8, 8)
        Me.chkSOCSketchAttached.Name = "chkSOCSketchAttached"
        Me.chkSOCSketchAttached.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkSOCSketchAttached.Size = New System.Drawing.Size(107, 15)
        Me.chkSOCSketchAttached.TabIndex = 46
        Me.chkSOCSketchAttached.Text = "Sketch Attached"
        Me.chkSOCSketchAttached.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlSOCDisplay
        '
        Me.pnlSOCDisplay.Controls.Add(Me.lblSOC)
        Me.pnlSOCDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlSOCDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlSOCDisplay.Name = "pnlSOCDisplay"
        Me.pnlSOCDisplay.Size = New System.Drawing.Size(864, 24)
        Me.pnlSOCDisplay.TabIndex = 43
        '
        'lblSOC
        '
        Me.lblSOC.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblSOC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblSOC.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSOC.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblSOC.Location = New System.Drawing.Point(0, 0)
        Me.lblSOC.Name = "lblSOC"
        Me.lblSOC.Size = New System.Drawing.Size(864, 24)
        Me.lblSOC.TabIndex = 2
        Me.lblSOC.Text = "SIGNIFICANT OPERATIONAL COMPLIANCE (SOC)"
        Me.lblSOC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CheckList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1028, 709)
        Me.Controls.Add(Me.pnlChecklistDetails)
        Me.Controls.Add(Me.pnlChecklistControls)
        Me.Controls.Add(Me.pnlChecklistBottom)
        Me.Name = "CheckList"
        Me.Text = "CheckList"
        Me.pnlChecklistBottom.ResumeLayout(False)
        Me.pnlChecklistControls.ResumeLayout(False)
        Me.pnlChecklistDetails.ResumeLayout(False)
        Me.pnlMW.ResumeLayout(False)
        CType(Me.ugMW, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMWAdd.ResumeLayout(False)
        Me.pnlMWDisplay.ResumeLayout(False)
        Me.pnlMaster.ResumeLayout(False)
        Me.pnlMasterDetails.ResumeLayout(False)
        Me.pnlContact.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.ugContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlContactBottom.ResumeLayout(False)
        Me.pnlContactTop.ResumeLayout(False)
        Me.pnlMasterDetailsGrid.ResumeLayout(False)
        CType(Me.ugInspector, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMasterDetailsTop.ResumeLayout(False)
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMasterDisplay.ResumeLayout(False)
        Me.pnlComments.ResumeLayout(False)
        Me.pnlCommentsDisplay.ResumeLayout(False)
        Me.pnlTOS.ResumeLayout(False)
        CType(Me.ugTOS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTOSDisplay.ResumeLayout(False)
        Me.pnlVisual.ResumeLayout(False)
        CType(Me.ugVisual, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlVisualDisplay.ResumeLayout(False)
        Me.pnlCatLeak.ResumeLayout(False)
        CType(Me.ugCatLeak, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCatLeakDisplay.ResumeLayout(False)
        Me.pnlPipeLeak.ResumeLayout(False)
        CType(Me.ugPipeLeak, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPipeLeakAdd.ResumeLayout(False)
        Me.pnlPipeLeakDisplay.ResumeLayout(False)
        Me.pnlTankLeak.ResumeLayout(False)
        CType(Me.ugTankLeak, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTankLeakAdd.ResumeLayout(False)
        Me.pnlTankLeakDisplay.ResumeLayout(False)
        Me.pnlCP.ResumeLayout(False)
        CType(Me.ugCP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCPAdd.ResumeLayout(False)
        Me.pnlCPDisplay.ResumeLayout(False)
        Me.pnlSpill.ResumeLayout(False)
        CType(Me.ugSpill, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlSpillDisplay.ResumeLayout(False)
        Me.pnlReg.ResumeLayout(False)
        Me.pnlRegDisplay.ResumeLayout(False)
        CType(Me.ugReg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlInspectionCitations.ResumeLayout(False)
        Me.pnlInspecCitationsDisplay.ResumeLayout(False)
        CType(Me.ugCitation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTanksPipes.ResumeLayout(False)
        Me.pnlTanksPipesDetails.ResumeLayout(False)
        Me.pnlTerminationsDetails.ResumeLayout(False)
        CType(Me.ugTerminations, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTerminationsTop.ResumeLayout(False)
        Me.pnlPipeDetails.ResumeLayout(False)
        CType(Me.ugPipes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPipeTop.ResumeLayout(False)
        Me.pnlTanksDetails.ResumeLayout(False)
        CType(Me.ugTanks, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTanksTop.ResumeLayout(False)
        Me.pnlTanksPipesDisplay.ResumeLayout(False)
        Me.pnlSOC.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.ugSOC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.pnlSOCDisplay.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Support Routines"
    Private Sub HideVisiblePanels(ByVal bolActivePanel As Boolean, ByRef pnlActivePanel As Panel)
        Try
            pnlMaster.Visible = False
            pnlTanksPipes.Visible = False
            pnlReg.Visible = False
            pnlInspectionCitations.Visible = False
            pnlSpill.Visible = False
            pnlCP.Visible = False
            pnlTankLeak.Visible = False
            pnlPipeLeak.Visible = False
            pnlCatLeak.Visible = False
            pnlVisual.Visible = False
            pnlTOS.Visible = False
            pnlComments.Visible = False
            pnlMW.Visible = False
            pnlSOC.Visible = False
            pnlActivePanel.Visible = bolActivePanel
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ToggleButtonAppearance(ByVal btnCurrentSelected As Button)
        Try
            If Not btnPrevSelected Is Nothing Then
                If btnPrevSelected.Name <> btnCurrentSelected.Name Then
                    btnPrevSelected.BackColor = System.Drawing.SystemColors.Control
                    btnPrevSelected.ForeColor = System.Drawing.SystemColors.ControlText
                End If
            End If
            btnCurrentSelected.BackColor = Color.Gray
            btnCurrentSelected.ForeColor = Color.White
            btnPrevSelected = btnCurrentSelected
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub LoadChecklistMaster()
        Try
            HideVisiblePanels(True, pnlMaster)

            ' facility address
            oInspection.CheckListMaster.Owner.Facilities.FacilityAddresses.Retrieve(oInspection.CheckListMaster.Owner.Facilities.AddressID)
            txtFacilityName.Text = oInspection.CheckListMaster.Owner.Facilities.Name
            txtAddress.Text = UIUtilsGen.FormatAddress(oInspection.CheckListMaster.Owner.Facilities.FacilityAddresses, True)

            txtLatitudeDegree.Text = IIf(oInspection.CheckListMaster.Owner.Facilities.LatitudeDegree = -1, String.Empty, oInspection.CheckListMaster.Owner.Facilities.LatitudeDegree)
            txtLatitudeMin.Text = IIf(oInspection.CheckListMaster.Owner.Facilities.LatitudeMinutes = -1, String.Empty, oInspection.CheckListMaster.Owner.Facilities.LatitudeMinutes)
            txtLatitudeSec.Text = IIf(oInspection.CheckListMaster.Owner.Facilities.LatitudeSeconds < 0, String.Empty, FormatNumber(oInspection.CheckListMaster.Owner.Facilities.LatitudeSeconds, 2, TriState.True, TriState.False, TriState.True))
            txtLongitudeDegree.Text = IIf(oInspection.CheckListMaster.Owner.Facilities.LongitudeDegree = -1, String.Empty, oInspection.CheckListMaster.Owner.Facilities.LongitudeDegree)
            txtLongitudeMin.Text = IIf(oInspection.CheckListMaster.Owner.Facilities.LongitudeMinutes = -1, String.Empty, oInspection.CheckListMaster.Owner.Facilities.LongitudeMinutes)
            txtLongitudeSec.Text = IIf(oInspection.CheckListMaster.Owner.Facilities.LongitudeSeconds < 0, String.Empty, FormatNumber(oInspection.CheckListMaster.Owner.Facilities.LongitudeSeconds, 2, TriState.True, TriState.False, TriState.True))
            If oInspection.CheckListMaster.Owner.PersonID = 0 And oInspection.CheckListMaster.Owner.OrganizationID <> 0 Then
                ' org
                txtUstOwner.Text = oInspection.CheckListMaster.Owner.Organization.Company
            ElseIf oInspection.CheckListMaster.Owner.PersonID <> 0 And oInspection.CheckListMaster.Owner.OrganizationID = 0 Then
                ' person
                txtUstOwner.Text = oInspection.CheckListMaster.Owner.Persona.FirstName.Trim + " " + oInspection.CheckListMaster.Owner.Persona.LastName.Trim
            End If
            txtOwnersAddress.Text = oInspection.CheckListMaster.Owner.Addresses.AddressLine1
            txtOwnerCity.Text = oInspection.CheckListMaster.Owner.Addresses.City
            txtOwnerState.Text = oInspection.CheckListMaster.Owner.Addresses.State
            txtOwnerZip.Text = oInspection.CheckListMaster.Owner.Addresses.Zip
            'txtOwnersPhone.Text = oInspection.CheckListMaster.Owner.PhoneNumberOne
            mskTxtOwnerPhone.SelText = IIf(oInspection.CheckListMaster.Owner.PhoneNumberOne.Length = 0, "", Trim(oInspection.CheckListMaster.Owner.PhoneNumberOne))
            txtOwnersRep.Text = oInspection.OwnersRep.Trim

            Dim strFees As String = String.Empty
            Dim ds As DataSet = oInspection.CheckListMaster.Owner.RunSQLQuery("SELECT dbo.udfGetOwnerPastDueFees(" + oInspection.OwnerID.ToString + ",0," + oInspection.FacilityID.ToString + ")")
            If ds.Tables(0).Rows(0)(0) > 0 Then
                strFees = ds.Tables(0).Rows(0)(0)
                strFees = strFees.Split(".")(0)
            Else
                strFees = "$0"
            End If
            txtFacility.Text = strFees

            strFees = String.Empty
            ds = oInspection.CheckListMaster.Owner.RunSQLQuery("SELECT dbo.udfGetOwnerPastDueFees(" + oInspection.OwnerID.ToString + ",0,NULL)")
            If ds.Tables(0).Rows(0)(0) > 0 Then
                strFees = ds.Tables(0).Rows(0)(0)
                strFees = strFees.Split(".")(0)
            Else
                strFees = "$0"
            End If
            txtOwner.Text = strFees

            ugInspector.DataSource = Nothing
            ugInspector.DataSource = oInspection.CheckListMaster.GetCLInspectionHistory(bolReadOnly)
            SetupInspector(bolReadOnly)

            chkCapCandidate.Checked = oInspection.CheckListMaster.Owner.Facilities.CAPCandidate

            ' 2879 lust
            ds = oInspection.CheckListMaster.Owner.RunSQLQuery("select distinct sm.[user_name] from tbltec_event t inner join tblsys_ust_staff_master sm on t.event_project_manager_id = sm.staff_id where t.deleted = 0 and t.event_status = 624 and t.facility_id = " + oInspection.FacilityID.ToString)
            lblLustActiveValue.Text = "NO"
            lblLustPMValue.Text = "N/A"
            If ds.Tables.Count > 0 Then
                If ds.Tables(0).Rows.Count > 0 Then
                    lblLustActiveValue.Text = "YES"
                    For Each dr As DataRow In ds.Tables(0).Rows
                        lblLustPMValue.Text = dr(0) + ", "
                    Next
                End If
            End If
            If lblLustPMValue.Text <> String.Empty Then
                lblLustPMValue.Text = lblLustPMValue.Text.Trim.TrimEnd(",")
            End If

            ShowHideMW()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub MakeFormReadOnly(ByVal val As Boolean)
        Try
            txtFacilityName.Enabled = Not val
            txtAddress.Enabled = Not val
            txtLatitudeDegree.Enabled = Not val
            txtLatitudeMin.Enabled = Not val
            txtLatitudeSec.Enabled = Not val
            txtLongitudeDegree.Enabled = Not val
            txtLongitudeMin.Enabled = Not val
            txtLongitudeSec.Enabled = Not val
            txtOwnersRep.Enabled = Not val
            chkCapCandidate.Enabled = Not val
            txtComments.ReadOnly = val
            chkSOCSketchAttached.Enabled = Not val

            btnSave.Enabled = Not val
            btnSubmitToCE.Enabled = Not val
            btnUnsubmit.Enabled = Not val
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SetupGrid(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim ugcol As Infragistics.Win.UltraWinGrid.UltraGridColumn
        Dim hasYesNo As Boolean = False

        Try
            ug.DisplayLayout.Bands(0).Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ug.DisplayLayout.Bands(0).Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
            ug.DisplayLayout.Bands(0).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed

            ug.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            ug.DisplayLayout.Override.CellAppearance.FontData.SizeInPoints = 7.5

            For Each ugcol In ug.DisplayLayout.Bands(0).Columns
                Select Case (ugcol.Key)
                    Case "CL_POSITION"
                        ug.DisplayLayout.Bands(0).Columns("CL_POSITION").Hidden = True

                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        ug.DisplayLayout.Bands(0).Columns("CL_POSITION").Header.VisiblePosition = 0
                        ug.DisplayLayout.Bands(0).Columns("CL_POSITION").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    Case "Line#"
                        ug.DisplayLayout.Bands(0).Columns("Line#").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ug.DisplayLayout.Bands(0).Columns("Line#").Width = 55
                        ug.DisplayLayout.Bands(0).Columns("Line#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                        ug.DisplayLayout.Bands(0).Columns("Line#").TabStop = False
                    Case "Question"
                        ug.DisplayLayout.Bands(0).Columns("Question").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ug.DisplayLayout.Bands(0).Columns("Question").Width = 545
                        ug.DisplayLayout.Bands(0).Columns("Question").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                        ug.DisplayLayout.Bands(0).Columns("Question").TabStop = False
                        'ug.DisplayLayout.Bands(0).Columns("Question").CellMultiLine = Infragistics.Win.DefaultableBoolean.True
                    Case "Citation"
                        ug.DisplayLayout.Bands(0).Columns("Citation").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ug.DisplayLayout.Bands(0).Columns("Citation").Width = 500
                        ug.DisplayLayout.Bands(0).Columns("Citation").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                        ug.DisplayLayout.Bands(0).Columns("Citation").TabStop = False
                    Case "CCAT"
                        ug.DisplayLayout.Bands(0).Columns("CCAT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ug.DisplayLayout.Bands(0).Columns("CCAT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                        ug.DisplayLayout.Bands(0).Columns("CCAT").TabStop = False
                        ug.DisplayLayout.Bands(0).Columns("CCAT").Width = 450

                    Case "Yes"
                        ug.DisplayLayout.Bands(0).Columns("Yes").Width = 50
                        ug.DisplayLayout.Bands(0).Columns("Yes").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                        If ug.Name <> ugSOC.Name Then
                            hasYesNo = True
                        End If
                    Case "No"
                        ug.DisplayLayout.Bands(0).Columns("No").Width = 50
                        ug.DisplayLayout.Bands(0).Columns("No").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                        If ug.Name <> ugSOC.Name Then
                            hasYesNo = True
                        End If
                        ug.DisplayLayout.Bands(0).Columns("No").TabStop = False
                    Case "ID"
                        ug.DisplayLayout.Bands(0).Columns("ID").Hidden = True
                    Case "INSPECTION_ID"
                        ug.DisplayLayout.Bands(0).Columns("INSPECTION_ID").Hidden = True
                    Case "QUESTION_ID"
                        ug.DisplayLayout.Bands(0).Columns("QUESTION_ID").Hidden = True
                    Case "SOC"
                        ug.DisplayLayout.Bands(0).Columns("SOC").Hidden = True
                    Case "RESPONSE"
                        ug.DisplayLayout.Bands(0).Columns("RESPONSE").Hidden = True
                    Case "HEADER"
                        ug.DisplayLayout.Bands(0).Columns("HEADER").Hidden = True
                    Case "CITATION"
                        ug.DisplayLayout.Bands(0).Columns("CITATION").Hidden = True
                    Case "DELETED"
                        ug.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
                    Case "Citations"
                        ug.DisplayLayout.Bands(0).Columns("Citations").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ug.DisplayLayout.Bands(0).Columns("Citations").TabStop = False
                        'ug.DisplayLayout.Bands(0).Columns("Citations").Width = 100
                    Case "FSOC_LK_PREVENT"
                        ug.DisplayLayout.Bands(0).Columns("FSOC_LK_PREVENT").Hidden = True
                    Case "FSOC_LK_PRE_CITATION"
                        ug.DisplayLayout.Bands(0).Columns("FSOC_LK_PRE_CITATION").Hidden = True
                    Case "FAC_SOC_LK_DETECTION"
                        ug.DisplayLayout.Bands(0).Columns("FAC_SOC_LK_DETECTION").Hidden = True
                    Case "FSOC_LK_DET_CITATION"
                        ug.DisplayLayout.Bands(0).Columns("FSOC_LK_DET_CITATION").Hidden = True
                    Case "FSOC_LK_PRE_LK_DET"
                        ug.DisplayLayout.Bands(0).Columns("FSOC_LK_PRE_LK_DET").Hidden = True
                    Case "Line Numbers"
                        ug.DisplayLayout.Bands(0).Columns("Line Numbers").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ug.DisplayLayout.Bands(0).Columns("Line Numbers").TabStop = False
                    Case "FORE_COLOR"
                        ug.DisplayLayout.Bands(0).Columns("FORE_COLOR").Hidden = True
                    Case "BACK_COLOR"
                        ug.DisplayLayout.Bands(0).Columns("BACK_COLOR").Hidden = True
                End Select
            Next

            If ug.DisplayLayout.Bands.Count > 1 Then
                ug.DisplayLayout.Bands(1).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                ug.DisplayLayout.Bands(1).Columns("ID").Header.VisiblePosition = 0
                ug.DisplayLayout.Bands(1).Columns("ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In ug.DisplayLayout.Bands(1).Columns
                    Select Case col.Key
                        Case "Line#"
                            ug.DisplayLayout.Bands(1).Columns("Line#").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                            'ug.DisplayLayout.Bands(1).Columns("Line#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Line#").TabStop = False
                            ug.DisplayLayout.Bands(1).Columns("Line#").Width = 60
                        Case "Well#"
                            ug.DisplayLayout.Bands(1).Columns("Well#").MaskInput = "nnnnnnnnn"
                            ug.DisplayLayout.Bands(1).Columns("Well#").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                            ug.DisplayLayout.Bands(1).Columns("Well#").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

                            'ug.DisplayLayout.Bands(1).Columns("Well Depth").MaskInput = "nnn,nnn,nnn"
                            'ug.DisplayLayout.Bands(1).Columns("Well Depth").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                            'ug.DisplayLayout.Bands(1).Columns("Well Depth").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

                            'ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Water").MaskInput = "nnn,nnn,nnn"
                            'ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Water").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                            'ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Water").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

                            'ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Slots").MaskInput = "nnn,nnn,nnn"
                            'ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Slots").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                            'ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Slots").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

                            ug.DisplayLayout.Bands(1).Columns("Well#").Width = 50
                            ug.DisplayLayout.Bands(1).Columns("Well Depth").Width = 70
                            ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Water").Width = 50
                            ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Slots").Width = 50
                            ug.DisplayLayout.Bands(1).ColHeaderLines = 2
                            ug.DisplayLayout.Bands(1).Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes
                            ug.DisplayLayout.Bands(1).Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True

                            ug.DisplayLayout.Bands(0).Columns("Question").ColSpan = 8
                            ug.DisplayLayout.Bands(0).Columns("CCAT").Width = 100

                            'ug.DisplayLayout.Bands(1).Columns("Well#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Well Depth").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Water").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Depth to" + vbCrLf + "Slots").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Surface Sealed" + vbCrLf + "Yes").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Surface Sealed" + vbCrLf + "No").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Well Caps" + vbCrLf + "Yes").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Well Caps" + vbCrLf + "No").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Inspector's Observations").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            'ug.DisplayLayout.Bands(1).Columns("ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                            'ug.DisplayLayout.Bands(1).Columns("QUESTION_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                            ug.DisplayLayout.Bands(1).SortedColumns.Clear()
                            ug.DisplayLayout.Bands(1).Columns("Well#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                            ug.DisplayLayout.Bands(1).Columns("LINE_NUMBER").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        Case "TANK_LINE"
                            ug.DisplayLayout.Bands(1).Columns("TANK_LINE").Hidden = True
                        Case "RECITIFIER_ON"
                            ug.DisplayLayout.Bands(1).Columns("RECITIFIER_ON").Hidden = True
                            'Case "INOP_HOW_LONG"
                            'ug.DisplayLayout.Bands(1).Columns("INOP_HOW_LONG").Hidden = True
                        Case "Amps"
                            ug.DisplayLayout.Bands(1).Columns("Volts").MaskInput = "nnn,nnn,nnn.nn"
                            ug.DisplayLayout.Bands(1).Columns("Volts").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                            ug.DisplayLayout.Bands(1).Columns("Volts").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

                            ug.DisplayLayout.Bands(1).Columns("Amps").MaskInput = "nnn,nnn,nnn.nn"
                            ug.DisplayLayout.Bands(1).Columns("Amps").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                            ug.DisplayLayout.Bands(1).Columns("Amps").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

                            ug.DisplayLayout.Bands(1).Columns("Hours").MaskInput = "nnn,nnn,nnn.nn"
                            ug.DisplayLayout.Bands(1).Columns("Hours").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                            ug.DisplayLayout.Bands(1).Columns("Hours").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

                            'ug.DisplayLayout.Bands(1).Columns("How Long").MaskInput = "nnn,nnn,nnn"
                            'ug.DisplayLayout.Bands(1).Columns("How Long").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                            'ug.DisplayLayout.Bands(1).Columns("How Long").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw
                            ug.DisplayLayout.Bands(1).Columns("How Long").Header.Caption = "How long has rectifier been turned off?"

                            ug.DisplayLayout.Bands(1).Columns("Amps").Width = 75
                            ug.DisplayLayout.Bands(1).Columns("Hours").Width = 75
                            ug.DisplayLayout.Bands(1).Columns("How Long").Width = 450
                            ug.DisplayLayout.Bands(0).Columns("Question").ColSpan = 2

                            ' sorting
                            ug.DisplayLayout.Bands(1).Columns("Amps").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Hours").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("Volts").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                            ug.DisplayLayout.Bands(1).Columns("How Long").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                        Case "SURFACE_SEALED"
                            ug.DisplayLayout.Bands(1).Columns("SURFACE_SEALED").Hidden = True
                        Case "WELL_CAPS"
                            ug.DisplayLayout.Bands(1).Columns("WELL_CAPS").Hidden = True
                        Case "CITATION"
                            ug.DisplayLayout.Bands(1).Columns("CITATION").Hidden = True
                        Case "LINE_NUMBER"
                            ug.DisplayLayout.Bands(1).Columns("LINE_NUMBER").Hidden = True
                    End Select
                Next

                If ug.Name = "ugCP" Then
                    ' populate columns
                    If ug.DisplayLayout.ValueLists.All.Length = 0 Then
                        ug.DisplayLayout.ValueLists.Add("TANK_NUM")
                        ug.DisplayLayout.ValueLists.Add("PIPE_NUM")
                        ug.DisplayLayout.ValueLists.Add("TERM_NUM")
                    End If

                    ' CP Tested by Inspector (Yes/No)
                    ug.DisplayLayout.Bands(2).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                    ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                    ug.DisplayLayout.Bands(2).Columns("ID").Hidden = True
                    ug.DisplayLayout.Bands(2).Columns("INSPECTION_ID").Hidden = True
                    ug.DisplayLayout.Bands(2).Columns("QUESTION_ID").Hidden = True
                    ug.DisplayLayout.Bands(2).Columns("TESTED_BY_INSPECTOR_RESPONSE").Hidden = True
                    ug.DisplayLayout.Bands(2).Columns("Yes").Width = 50
                    ug.DisplayLayout.Bands(2).Columns("No").Width = 50
                    ug.DisplayLayout.Bands(2).Columns("BLANK").Header.Caption = ""
                    ug.DisplayLayout.Bands(2).Columns("BLANK").Width = 575
                    ug.DisplayLayout.Bands(2).Columns("BLANK").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled

                    ug.DisplayLayout.Bands(3).Hidden = False
                    ug.DisplayLayout.Bands(4).Hidden = False
                    ug.DisplayLayout.Bands(5).Hidden = False
                    If Not hasCPTank Then
                        ug.DisplayLayout.Bands(3).Hidden = True
                        ug.DisplayLayout.Bands(4).Hidden = True
                        ug.DisplayLayout.Bands(5).Hidden = True
                    Else
                        ' Galvanic Impressed Current
                        ug.DisplayLayout.Bands(3).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        ug.DisplayLayout.Bands(3).Columns("ID").Hidden = True
                        ug.DisplayLayout.Bands(3).Columns("INSPECTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(3).Columns("QUESTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(3).Columns("GALVANIC_IC_RESPONSE").Hidden = True
                        ug.DisplayLayout.Bands(3).Hidden = True

                        ' Description of Remote Reference Cell Placement
                        ug.DisplayLayout.Bands(4).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        ug.DisplayLayout.Bands(4).Columns("ID").Hidden = True
                        ug.DisplayLayout.Bands(4).Columns("INSPECTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(4).Columns("QUESTION_ID").Hidden = True

                        ' CP Readings
                        ' sorting
                        ug.DisplayLayout.Bands(5).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        'ug.DisplayLayout.Bands(5).Columns("Tank#").Header.VisiblePosition = 0
                        ug.DisplayLayout.Bands(5).SortedColumns.Clear()
                        ug.DisplayLayout.Bands(5).Columns("TANK_INDEX").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(5).Columns("LINE_NUMBER").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(5).Columns("Tank#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(5).Columns("Tank#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled

                        pnlCPAdd.Visible = True
                        btnAddTankCP.Enabled = True
                        ug.DisplayLayout.Bands(5).Columns("Tank#").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                        ' populate the whole column as the table is the same for each row
                        ' Tank#
                        If ug.DisplayLayout.Bands(5).Columns("Tank#").ValueList Is Nothing Then
                            vListTankNum = New Infragistics.Win.ValueList
                            For i As Integer = 0 To oInspection.CheckListMaster.CPReadingTankIDs.Count - 1
                                vListTankNum.ValueListItems.Add(oInspection.CheckListMaster.CPReadingTankIDs.GetKey(i), oInspection.CheckListMaster.CPReadingTankIDs.GetByIndex(i))
                            Next
                            ug.DisplayLayout.Bands(5).Columns("Tank#").ValueList = vListTankNum
                        End If
                        For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In ug.DisplayLayout.Bands(5).Columns
                            Select Case col.Key
                                Case "Line#"
                                    ug.DisplayLayout.Bands(5).Columns("Line#").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                                    'ug.DisplayLayout.Bands(5).Columns("Line#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    ug.DisplayLayout.Bands(5).Columns("Line#").TabStop = False
                                    ug.DisplayLayout.Bands(5).Columns("Line#").Width = 60
                                Case "Fuel Type"
                                    ug.DisplayLayout.Bands(5).Columns("Fuel Type").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                                    ug.DisplayLayout.Bands(5).Columns("Fuel Type").TabStop = False
                                    ug.DisplayLayout.Bands(5).Columns("Fuel Type").Width = 60
                                Case "Tank#"
                                    ug.DisplayLayout.Bands(5).Columns("Tank#").Width = 50
                                    ug.DisplayLayout.Bands(5).Columns("Contact Point").Width = 100
                                    ug.DisplayLayout.Bands(5).Columns("Local Reference Cell Placement").Width = 100
                                    ug.DisplayLayout.Bands(5).Columns("Local/On").Width = 100
                                    ug.DisplayLayout.Bands(5).Columns("Remote/Off").Width = 100

                                    ug.DisplayLayout.Bands(5).Columns("Contact Point").Header.Caption = "Contact Point"
                                    ug.DisplayLayout.Bands(5).Columns("Local Reference Cell Placement").Header.Caption = "Local Reference" + vbCrLf + "Cell Placement"
                                    ug.DisplayLayout.Bands(5).ColHeaderLines = 2

                                    'ug.DisplayLayout.Bands(5).Columns("Pass").Width = 50
                                    'ug.DisplayLayout.Bands(5).Columns("Fail").Width = 50

                                    ug.DisplayLayout.Bands(5).Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes
                                    ug.DisplayLayout.Bands(5).Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True

                                    ug.DisplayLayout.Bands(0).Columns("Question").ColSpan = 5
                                    ug.DisplayLayout.Bands(0).Columns("Yes").Width = 50
                                    ug.DisplayLayout.Bands(0).Columns("No").Width = 50
                                    ug.DisplayLayout.Bands(0).Columns("CCAT").Width = 100

                                    'ug.DisplayLayout.Bands(5).Columns("Tank#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(5).Columns("Contact Point").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(5).Columns("Local Reference Cell Placement").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(5).Columns("Local/On").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(5).Columns("Remote/Off").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(5).Columns("Pass").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(5).Columns("Fail").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(5).Columns("Incon").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                Case "LINE_NUMBER"
                                    ug.DisplayLayout.Bands(5).Columns("LINE_NUMBER").Hidden = True
                                Case "ID"
                                    ug.DisplayLayout.Bands(5).Columns("ID").Hidden = True
                                Case "INSPECTION_ID"
                                    ug.DisplayLayout.Bands(5).Columns("INSPECTION_ID").Hidden = True
                                Case "QUESTION_ID"
                                    ug.DisplayLayout.Bands(5).Columns("QUESTION_ID").Hidden = True
                                Case "TANK_PIPE_ID"
                                    ug.DisplayLayout.Bands(5).Columns("TANK_PIPE_ID").Hidden = True
                                Case "TANK_PIPE_ENTITY_ID"
                                    ug.DisplayLayout.Bands(5).Columns("TANK_PIPE_ENTITY_ID").Hidden = True
                                Case "TANK_DISPENSER"
                                    ug.DisplayLayout.Bands(5).Columns("TANK_DISPENSER").Hidden = True
                                Case "GALVANIC"
                                    ug.DisplayLayout.Bands(5).Columns("GALVANIC").Hidden = True
                                Case "IMPRESSED_CURRENT"
                                    ug.DisplayLayout.Bands(5).Columns("IMPRESSED_CURRENT").Hidden = True
                                Case "PASSFAILINCON"
                                    ug.DisplayLayout.Bands(5).Columns("PASSFAILINCON").Hidden = True
                                Case "CITATION"
                                    ug.DisplayLayout.Bands(5).Columns("CITATION").Hidden = True
                                Case "Question"
                                    ug.DisplayLayout.Bands(5).Columns("Question").Hidden = True
                                Case "TANK_INDEX"
                                    ug.DisplayLayout.Bands(5).Columns("TANK_INDEX").Hidden = True
                            End Select
                        Next
                    End If

                    ' CP Tested by Inspector (Yes/No)
                    ug.DisplayLayout.Bands(6).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                    ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                    ug.DisplayLayout.Bands(6).Columns("ID").Hidden = True
                    ug.DisplayLayout.Bands(6).Columns("INSPECTION_ID").Hidden = True
                    ug.DisplayLayout.Bands(6).Columns("QUESTION_ID").Hidden = True
                    ug.DisplayLayout.Bands(6).Columns("TESTED_BY_INSPECTOR_RESPONSE").Hidden = True
                    ug.DisplayLayout.Bands(6).Columns("Yes").Width = 50
                    ug.DisplayLayout.Bands(6).Columns("No").Width = 50
                    ug.DisplayLayout.Bands(6).Columns("BLANK").Header.Caption = ""
                    ug.DisplayLayout.Bands(6).Columns("BLANK").Width = 575
                    ug.DisplayLayout.Bands(6).Columns("BLANK").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled

                    ug.DisplayLayout.Bands(7).Hidden = False
                    ug.DisplayLayout.Bands(8).Hidden = False
                    ug.DisplayLayout.Bands(9).Hidden = False
                    If Not hasCPPipe Then
                        ug.DisplayLayout.Bands(7).Hidden = True
                        ug.DisplayLayout.Bands(8).Hidden = True
                        ug.DisplayLayout.Bands(9).Hidden = True
                    Else
                        ' Galvanic Impressed Current
                        ug.DisplayLayout.Bands(7).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        ug.DisplayLayout.Bands(7).Columns("ID").Hidden = True
                        ug.DisplayLayout.Bands(7).Columns("INSPECTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(7).Columns("QUESTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(7).Columns("GALVANIC_IC_RESPONSE").Hidden = True
                        ug.DisplayLayout.Bands(7).Hidden = True

                        ' Description of Remote Reference Cell Placement
                        ug.DisplayLayout.Bands(8).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        ug.DisplayLayout.Bands(8).Columns("ID").Hidden = True
                        ug.DisplayLayout.Bands(8).Columns("INSPECTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(8).Columns("QUESTION_ID").Hidden = True

                        ' CP Readings
                        ' sorting
                        ug.DisplayLayout.Bands(9).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        'ug.DisplayLayout.Bands(9).Columns("Pipe#").Header.VisiblePosition = 0
                        ug.DisplayLayout.Bands(9).SortedColumns.Clear()
                        ug.DisplayLayout.Bands(9).Columns("PIPE_INDEX").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(9).Columns("LINE_NUMBER").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(9).Columns("Pipe#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(9).Columns("Pipe#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled

                        pnlCPAdd.Visible = True
                        btnAddPipeCP.Enabled = True
                        ug.DisplayLayout.Bands(9).Columns("Pipe#").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                        ' populate the whole column as the table is the same for each row
                        ' Tank#
                        If ug.DisplayLayout.Bands(9).Columns("Pipe#").ValueList Is Nothing Then
                            vListPipeNum = New Infragistics.Win.ValueList
                            For i As Integer = 0 To oInspection.CheckListMaster.CPReadingPipeIDs.Count - 1
                                'vListPipeNum.ValueListItems.Add(oInspection.CheckListMaster.CPReadingPipeIDs.GetByIndex(i))
                                vListPipeNum.ValueListItems.Add(oInspection.CheckListMaster.CPReadingPipeIDs.GetKey(i), oInspection.CheckListMaster.CPReadingPipeIDs.GetByIndex(i))
                            Next
                            ug.DisplayLayout.Bands(9).Columns("Pipe#").ValueList = vListPipeNum
                        End If
                        For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In ug.DisplayLayout.Bands(9).Columns
                            Select Case col.Key
                                Case "Line#"
                                    ug.DisplayLayout.Bands(9).Columns("Line#").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                                    'ug.DisplayLayout.Bands(9).Columns("Line#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    ug.DisplayLayout.Bands(9).Columns("Line#").TabStop = False
                                    ug.DisplayLayout.Bands(9).Columns("Line#").Width = 60
                                Case "Fuel Type"
                                    ug.DisplayLayout.Bands(9).Columns("Fuel Type").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                                    ug.DisplayLayout.Bands(9).Columns("Fuel Type").TabStop = False
                                    ug.DisplayLayout.Bands(9).Columns("Fuel Type").Width = 60
                                Case "Pipe#"
                                    ug.DisplayLayout.Bands(9).Columns("Pipe#").Width = 50
                                    ug.DisplayLayout.Bands(9).Columns("Contact Point").Width = 100
                                    ug.DisplayLayout.Bands(9).Columns("Local Reference Cell Placement").Width = 100
                                    ug.DisplayLayout.Bands(9).Columns("Local/On").Width = 100
                                    ug.DisplayLayout.Bands(9).Columns("Remote/Off").Width = 100

                                    ug.DisplayLayout.Bands(9).Columns("Contact Point").Header.Caption = "Contact Point"
                                    ug.DisplayLayout.Bands(9).Columns("Local Reference Cell Placement").Header.Caption = "Local Reference" + vbCrLf + "Cell Placement"
                                    ug.DisplayLayout.Bands(9).ColHeaderLines = 2

                                    ug.DisplayLayout.Bands(9).Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes
                                    ug.DisplayLayout.Bands(9).Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True

                                    ug.DisplayLayout.Bands(0).Columns("Question").ColSpan = 5
                                    ug.DisplayLayout.Bands(0).Columns("Yes").Width = 50
                                    ug.DisplayLayout.Bands(0).Columns("No").Width = 50
                                    ug.DisplayLayout.Bands(0).Columns("CCAT").Width = 100

                                    'ug.DisplayLayout.Bands(9).Columns("Pipe#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(9).Columns("Contact Point").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(9).Columns("Local Reference Cell Placement").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(9).Columns("Local/On").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(9).Columns("Remote/Off").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(9).Columns("Pass").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(9).Columns("Fail").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(9).Columns("Incon").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                Case "LINE_NUMBER"
                                    ug.DisplayLayout.Bands(9).Columns("LINE_NUMBER").Hidden = True
                                Case "ID"
                                    ug.DisplayLayout.Bands(9).Columns("ID").Hidden = True
                                Case "INSPECTION_ID"
                                    ug.DisplayLayout.Bands(9).Columns("INSPECTION_ID").Hidden = True
                                Case "QUESTION_ID"
                                    ug.DisplayLayout.Bands(9).Columns("QUESTION_ID").Hidden = True
                                Case "TANK_PIPE_ID"
                                    ug.DisplayLayout.Bands(9).Columns("TANK_PIPE_ID").Hidden = True
                                Case "TANK_PIPE_ENTITY_ID"
                                    ug.DisplayLayout.Bands(9).Columns("TANK_PIPE_ENTITY_ID").Hidden = True
                                Case "TANK_DISPENSER"
                                    ug.DisplayLayout.Bands(9).Columns("TANK_DISPENSER").Hidden = True
                                Case "GALVANIC"
                                    ug.DisplayLayout.Bands(9).Columns("GALVANIC").Hidden = True
                                Case "IMPRESSED_CURRENT"
                                    ug.DisplayLayout.Bands(9).Columns("IMPRESSED_CURRENT").Hidden = True
                                Case "PASSFAILINCON"
                                    ug.DisplayLayout.Bands(9).Columns("PASSFAILINCON").Hidden = True
                                Case "CITATION"
                                    ug.DisplayLayout.Bands(9).Columns("CITATION").Hidden = True
                                Case "Question"
                                    ug.DisplayLayout.Bands(9).Columns("Question").Hidden = True
                                Case "PIPE_INDEX"
                                    ug.DisplayLayout.Bands(9).Columns("PIPE_INDEX").Hidden = True
                            End Select
                        Next
                    End If

                    ' CP Tested by Inspector (Yes/No)
                    ug.DisplayLayout.Bands(10).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                    ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                    ug.DisplayLayout.Bands(10).Columns("ID").Hidden = True
                    ug.DisplayLayout.Bands(10).Columns("INSPECTION_ID").Hidden = True
                    ug.DisplayLayout.Bands(10).Columns("QUESTION_ID").Hidden = True
                    ug.DisplayLayout.Bands(10).Columns("TESTED_BY_INSPECTOR_RESPONSE").Hidden = True
                    ug.DisplayLayout.Bands(10).Columns("Yes").Width = 50
                    ug.DisplayLayout.Bands(10).Columns("No").Width = 50
                    ug.DisplayLayout.Bands(10).Columns("BLANK").Header.Caption = ""
                    ug.DisplayLayout.Bands(10).Columns("BLANK").Width = 575
                    ug.DisplayLayout.Bands(10).Columns("BLANK").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled

                    ug.DisplayLayout.Bands(11).Hidden = False
                    ug.DisplayLayout.Bands(12).Hidden = False
                    ug.DisplayLayout.Bands(13).Hidden = False
                    If Not hasCPTerm Then
                        ug.DisplayLayout.Bands(11).Hidden = True
                        ug.DisplayLayout.Bands(12).Hidden = True
                        ug.DisplayLayout.Bands(13).Hidden = True
                    Else
                        ' Galvanic Impressed Current
                        ug.DisplayLayout.Bands(11).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        ug.DisplayLayout.Bands(11).Columns("ID").Hidden = True
                        ug.DisplayLayout.Bands(11).Columns("INSPECTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(11).Columns("QUESTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(11).Columns("GALVANIC_IC_RESPONSE").Hidden = True
                        ug.DisplayLayout.Bands(11).Hidden = True

                        ' Description of Remote Reference Cell Placement
                        ug.DisplayLayout.Bands(12).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        ug.DisplayLayout.Bands(12).Columns("ID").Hidden = True
                        ug.DisplayLayout.Bands(12).Columns("INSPECTION_ID").Hidden = True
                        ug.DisplayLayout.Bands(12).Columns("QUESTION_ID").Hidden = True

                        ' CP Readings
                        ' sorting
                        ug.DisplayLayout.Bands(13).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
                        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
                        'ug.DisplayLayout.Bands(13).Columns("Term#").Header.VisiblePosition = 0
                        ug.DisplayLayout.Bands(13).SortedColumns.Clear()
                        ug.DisplayLayout.Bands(13).Columns("TERM_INDEX").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(13).Columns("LINE_NUMBER").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(13).Columns("Term#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(13).Columns("Term#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled

                        pnlCPAdd.Visible = True
                        btnAddTermCP.Enabled = True
                        ug.DisplayLayout.Bands(13).Columns("Term#").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                        ' populate the whole column as the table is the same for each row
                        ' Term#
                        If ug.DisplayLayout.Bands(13).Columns("Term#").ValueList Is Nothing Then
                            vListTermNum = New Infragistics.Win.ValueList
                            For i As Integer = 0 To oInspection.CheckListMaster.CPReadingTermIDs.Count - 1
                                'vListTermNum.ValueListItems.Add(oInspection.CheckListMaster.CPReadingTermIDs.GetByIndex(i))
                                vListTermNum.ValueListItems.Add(oInspection.CheckListMaster.CPReadingTermIDs.GetKey(i), oInspection.CheckListMaster.CPReadingTermIDs.GetByIndex(i))
                            Next
                            ug.DisplayLayout.Bands(13).Columns("Term#").ValueList = vListTermNum
                        End If
                        For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In ug.DisplayLayout.Bands(13).Columns
                            Select Case col.Key
                                Case "Line#"
                                    ug.DisplayLayout.Bands(13).Columns("Line#").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                                    'ug.DisplayLayout.Bands(13).Columns("Line#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    ug.DisplayLayout.Bands(13).Columns("Line#").TabStop = False
                                    ug.DisplayLayout.Bands(13).Columns("Line#").Width = 60
                                Case "Fuel Type"
                                    ug.DisplayLayout.Bands(13).Columns("Fuel Type").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                                    ug.DisplayLayout.Bands(13).Columns("Fuel Type").TabStop = False
                                    ug.DisplayLayout.Bands(13).Columns("Fuel Type").Width = 60
                                Case "Term#"
                                    ug.DisplayLayout.Bands(13).Columns("Term#").Width = 50
                                    ug.DisplayLayout.Bands(13).Columns("Contact Point").Width = 100
                                    ug.DisplayLayout.Bands(13).Columns("Local Reference Cell Placement").Width = 100
                                    ug.DisplayLayout.Bands(13).Columns("Local/On").Width = 100
                                    ug.DisplayLayout.Bands(13).Columns("Remote/Off").Width = 100

                                    ug.DisplayLayout.Bands(13).Columns("Contact Point").Header.Caption = "Contact Point"
                                    ug.DisplayLayout.Bands(13).Columns("Local Reference Cell Placement").Header.Caption = "Local Reference" + vbCrLf + "Cell Placement"
                                    ug.DisplayLayout.Bands(13).ColHeaderLines = 2

                                    ug.DisplayLayout.Bands(13).Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.Yes
                                    ug.DisplayLayout.Bands(13).Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True

                                    ug.DisplayLayout.Bands(0).Columns("Question").ColSpan = 5
                                    ug.DisplayLayout.Bands(0).Columns("Yes").Width = 50
                                    ug.DisplayLayout.Bands(0).Columns("No").Width = 50
                                    ug.DisplayLayout.Bands(0).Columns("CCAT").Width = 200

                                    'ug.DisplayLayout.Bands(13).Columns("Term#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(13).Columns("Contact Point").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(13).Columns("Local Reference Cell Placement").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(13).Columns("Local/On").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(13).Columns("Remote/Off").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(13).Columns("Pass").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(13).Columns("Fail").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                    'ug.DisplayLayout.Bands(13).Columns("Incon").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Disabled
                                Case "LINE_NUMBER"
                                    ug.DisplayLayout.Bands(13).Columns("LINE_NUMBER").Hidden = True
                                Case "ID"
                                    ug.DisplayLayout.Bands(13).Columns("ID").Hidden = True
                                Case "INSPECTION_ID"
                                    ug.DisplayLayout.Bands(13).Columns("INSPECTION_ID").Hidden = True
                                Case "QUESTION_ID"
                                    ug.DisplayLayout.Bands(13).Columns("QUESTION_ID").Hidden = True
                                Case "TANK_PIPE_ID"
                                    ug.DisplayLayout.Bands(13).Columns("TANK_PIPE_ID").Hidden = True
                                Case "TANK_PIPE_ENTITY_ID"
                                    ug.DisplayLayout.Bands(13).Columns("TANK_PIPE_ENTITY_ID").Hidden = True
                                Case "TANK_DISPENSER"
                                    ug.DisplayLayout.Bands(13).Columns("TANK_DISPENSER").Hidden = True
                                Case "GALVANIC"
                                    ug.DisplayLayout.Bands(13).Columns("GALVANIC").Hidden = True
                                Case "IMPRESSED_CURRENT"
                                    ug.DisplayLayout.Bands(13).Columns("IMPRESSED_CURRENT").Hidden = True
                                Case "PASSFAILINCON"
                                    ug.DisplayLayout.Bands(13).Columns("PASSFAILINCON").Hidden = True
                                Case "CITATION"
                                    ug.DisplayLayout.Bands(13).Columns("CITATION").Hidden = True
                                Case "Question"
                                    ug.DisplayLayout.Bands(13).Columns("Question").Hidden = True
                                Case "TERM_INDEX"
                                    ug.DisplayLayout.Bands(13).Columns("TERM_INDEX").Hidden = True
                            End Select
                        Next
                    End If
                    pnlCPAdd.Visible = IIf(bolReadOnly, False, pnlCPAdd.Visible)
                End If

                ug.DisplayLayout.Bands(1).Columns("ID").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("INSPECTION_ID").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("QUESTION_ID").Hidden = True
            End If

            If hasYesNo Then
                Dim dr, drChild As Infragistics.Win.UltraWinGrid.UltraGridRow
                Dim a As New ColorConverter
                For Each dr In ug.Rows ' band 0
                    dr.Appearance.ForeColor = a.ConvertFromString(dr.Cells("FORE_COLOR").Text)
                    dr.Appearance.BackColor = a.ConvertFromString(dr.Cells("BACK_COLOR").Text)
                    If dr.Cells("HEADER").Value = True Then
                        dr.Cells("Yes").Hidden = True
                        dr.Cells("No").Hidden = True
                    End If

                    ' ugCP
                    If ug.Name = ugCP.Name And Not bolReadOnly Then
                        If dr.Cells("Line#").Value = "3.4" Then
                            If Not dr.ChildBands Is Nothing Then
                                If Not dr.ChildBands(0).Rows Is Nothing Then
                                    For Each drChild In dr.ChildBands(0).Rows
                                        SetupRectifierRow(drChild)
                                    Next
                                End If
                            End If
                        End If
                    ElseIf ug.Name = ugTankLeak.Name Or ug.Name = ugPipeLeak.Name Or ug.Name = ugMW.Name Then
                        ' ugTankLeak / ugPipeLeak
                        If dr.Cells("Line#").Value = "4.2.8" Or dr.Cells("Line#").Value = "5.2.8" Or dr.Cells("Line#").Value = "11" Then
                            dr.Expanded = True
                            If Not bolReadOnly Then
                                For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In dr.ChildBands(0).Rows
                                    SetupTankPipeLeakRow(ugChildRow)
                                Next
                            End If
                            If ug.Name = ugTankLeak.Name Then
                                pnlTankLeakAdd.Visible = True
                                btnAddTankMW.Enabled = True
                            ElseIf ug.Name = ugPipeLeak.Name Then
                                pnlPipeLeakAdd.Visible = True
                                btnAddPipeMW.Enabled = True
                            ElseIf ug.Name = ugMW.Name Then
                                dr.Cells("Line#").Hidden = True
                                pnlMWAdd.Visible = True
                                btnMWAdd.Enabled = True
                            End If
                        Else
                            dr.Expanded = False
                        End If
                    End If

                Next
            End If
            'If Not ug.DataSource Is Nothing Then
            '    If ug.Rows.Count > 0 Then
            '        ug.ActiveRow = ug.Rows(0)
            '    End If
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SetupInspector(Optional ByVal [readOnly] As Boolean = False)
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugInspector.DisplayLayout.AutoFitColumns = True
            ugInspector.Rows.ExpandAll(True)

            ugInspector.DisplayLayout.Bands(0).Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom
            ugInspector.DisplayLayout.Bands(0).Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True
            ugInspector.DisplayLayout.Bands(0).Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
            ugInspector.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            ugInspector.DisplayLayout.Bands(0).SortedColumns.Clear()
            ugInspector.DisplayLayout.Bands(0).Columns("DATE INSPECTED").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugInspector.DisplayLayout.Bands(0).Columns("TIME IN").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugInspector.DisplayLayout.Bands(0).Columns("TIME OUT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            ugInspector.DisplayLayout.Bands(0).Columns("INS_DATES_ID").Hidden = True
            ugInspector.DisplayLayout.Bands(0).Columns("INSPECTION_ID").Hidden = True
            ugInspector.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True

            ' specifying which cols are drop down
            ugInspector.DisplayLayout.Bands(0).Columns("DEQ INSPECTOR").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            ugInspector.DisplayLayout.Bands(0).Columns("TIME IN").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            ugInspector.DisplayLayout.Bands(0).Columns("TIME OUT").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

            ' populate columns
            If ugInspector.DisplayLayout.ValueLists.All.Length = 0 Then
                ugInspector.DisplayLayout.ValueLists.Add("DEQ_INSPECTOR")
                ugInspector.DisplayLayout.ValueLists.Add("TIME_IN")
                ugInspector.DisplayLayout.ValueLists.Add("TIME_OUT")
            End If

            ' populate the whole column as the table is the same for each row
            If vListInspectors Is Nothing Then
                vListInspectors = New Infragistics.Win.ValueList
                For Each row As DataRow In oInspection.GetInspectors.Tables(0).Rows
                    vListInspectors.ValueListItems.Add(row.Item("STAFF_ID"), row.Item("USER_NAME").ToString)
                Next
            End If

            If vListTime Is Nothing Then
                vListTime = New Infragistics.Win.ValueList
                For Each row As DataRow In oInspection.GetInspectionTimes.Tables(0).Rows
                    vListTime.ValueListItems.Add(row.Item("PROPERTY_NAME"), row.Item("PROPERTY_NAME").ToString)
                Next
            End If

            ' DEQ INSPECTOR
            If ugInspector.DisplayLayout.Bands(0).Columns("DEQ INSPECTOR").ValueList Is Nothing Then
                ugInspector.DisplayLayout.Bands(0).Columns("DEQ INSPECTOR").ValueList = vListInspectors
            End If
            ' TIME IN
            If ugInspector.DisplayLayout.Bands(0).Columns("TIME IN").ValueList Is Nothing Then
                ugInspector.DisplayLayout.Bands(0).Columns("TIME IN").ValueList = vListTime
            End If
            ' TIME OUT
            If ugInspector.DisplayLayout.Bands(0).Columns("TIME OUT").ValueList Is Nothing Then
                ugInspector.DisplayLayout.Bands(0).Columns("TIME OUT").ValueList = vListTime
            End If

            If Not [readOnly] Then
                For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugInspector.Rows
                    If vListInspectors.FindByDataValue(ugRow.Cells("DEQ INSPECTOR").Value) Is Nothing Then
                        ugRow.Cells("DEQ INSPECTOR").Value = DBNull.Value
                    End If
                    If vListTime.FindByDataValue(ugRow.Cells("TIME IN").Text) Is Nothing Then
                        ugRow.Cells("TIME IN").Value = DBNull.Value
                    End If
                    If vListTime.FindByDataValue(ugRow.Cells("TIME OUT").Text) Is Nothing Then
                        ugRow.Cells("TIME OUT").Value = DBNull.Value
                    End If
                Next
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PrepTankPipeTerm()
        Dim dsSet As DataSet
        Dim dsSet2 As DataSet
        Dim dsSet3 As DataSet
        Try
            dsSet = oInspection.CheckListMaster.TanksPipesTables(oInspection.FacilityID)
            ToggleButtonAppearance(btnTanksPipes)
            HideVisiblePanels(True, pnlTanksPipes)
            Me.ugTanks.DataSource = dsSet.Tables(0)

            dsSet2 = dsSet.Copy
            dsSet3 = dsSet.Copy

            dsSet2.Relations.RemoveAt(1)
            dsSet2.Tables.RemoveAt(0)
            dsSet2.Tables.RemoveAt(1)
            dsSet2.Tables.RemoveAt(2)

            dsSet3.Relations.RemoveAt(0)
            dsSet3.Tables.RemoveAt(0)
            dsSet3.Tables.RemoveAt(0)
            dsSet3.Tables.RemoveAt(1)



            Me.ugPipes.DataSource = dsSet2
            Me.ugTerminations.DataSource = dsSet3
            ugTanks.DrawFilter = rp
            ugPipes.DrawFilter = rp
            ugTerminations.DrawFilter = rp
            SetupTankPipeTerm(TankPipeTermGrid.All, bolReadOnly)


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub InitSetupTankPipeTermVariables()
    '    CPTypeCount = 0
    '    CPInstalledCount = 0
    '    CPTestedCount = 0
    '    LeakDetectionCount = 0
    '    OverFillCount = 0
    '    LinedCount = 0
    '    LiningInspectedCount = 0
    '    PTTCount = 0
    '    BrandCount = 0
    '    PriLeakDetectionCount = 0
    '    SecLeakDetectionCount = 0
    '    ALLDTestedCount = 0
    '    TermCPTestedCount = 0
    '    LastUsedCount = 0
    '    PipeLastUsedCount = 0
    '    TermDispCPTypeCount = 0
    '    TermTankCPTypeCount = 0
    'End Sub
    Private Sub SetupTankPipeTerm(Optional ByVal grid As TankPipeTermGrid = TankPipeTermGrid.All, Optional ByVal [readOnly] As Boolean = False)
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If grid = TankPipeTermGrid.Tank Or grid = TankPipeTermGrid.All Then
                '************************************************************************
                ' Tank
                '************************************************************************

                If oTank Is Nothing Then
                    oTank = New MUSTER.BusinessLogic.pTank
                End If

                'InitSetupTankPipeTermVariables()

                ugTanks.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free

                ugTanks.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
                ugTanks.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No

                For g As Integer = 0 To 0
                    ugTanks.DisplayLayout.Bands(g).Columns("Tank #").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ugTanks.DisplayLayout.Bands(g).Columns("STATUS").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ugTanks.DisplayLayout.Bands(g).Columns("COMPARTMENT #").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                    ugTanks.DisplayLayout.Bands(g).Columns("FACILITY_ID").Hidden = True
                    ugTanks.DisplayLayout.Bands(g).Columns("TANK_ID").Hidden = True
                    ugTanks.DisplayLayout.Bands(g).Columns("COMPARTMENT").Hidden = True
                    ugTanks.DisplayLayout.Bands(g).Columns("COMPARTMENT_NUMBER").Hidden = True
                    ugTanks.DisplayLayout.Bands(g).Columns("POSITION").Hidden = True

                    'SIZE
                    ugTanks.DisplayLayout.Bands(g).Columns("SIZE").MaskInput = "nnn,nnn,nnn"
                    ugTanks.DisplayLayout.Bands(g).Columns("SIZE").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
                    ugTanks.DisplayLayout.Bands(g).Columns("SIZE").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

                    ' specifying which cols are drop down
                    ugTanks.DisplayLayout.Bands(g).Columns("CONTENTS").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTanks.DisplayLayout.Bands(g).Columns("FUEL TYPE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTanks.DisplayLayout.Bands(g).Columns("MATERIALS").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTanks.DisplayLayout.Bands(g).Columns("OPTIONS").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTanks.DisplayLayout.Bands(g).Columns("CP TYPE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTanks.DisplayLayout.Bands(g).Columns("LEAK DETECTION").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTanks.DisplayLayout.Bands(g).Columns("OVERFILL TYPE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

                Next g

                ' populate columns
                If ugTanks.DisplayLayout.ValueLists.All.Length = 0 Then
                    ugTanks.DisplayLayout.ValueLists.Add("CONTENTS")
                    ugTanks.DisplayLayout.ValueLists.Add("FUEL_TYPE")
                    ugTanks.DisplayLayout.ValueLists.Add("MATERIAL_OF_CONSTRUCTION")
                    ugTanks.DisplayLayout.ValueLists.Add("CONSTRUCTION_OPTIONS")
                    ugTanks.DisplayLayout.ValueLists.Add("CP_TYPE")
                    ugTanks.DisplayLayout.ValueLists.Add("LEAK_DETECTION")
                    ugTanks.DisplayLayout.ValueLists.Add("OVER_FILL")
                End If

                For g As Integer = 0 To 0
                    ' populate the whole column as the table is the same for each row
                    ' Contents
                    If ugTanks.DisplayLayout.Bands(g).Columns("CONTENTS").ValueList Is Nothing Then
                        vListCompSubstance = New Infragistics.Win.ValueList
                        For Each row As DataRow In oTank.PopulateCompartmentSubstance.Rows
                            vListCompSubstance.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugTanks.DisplayLayout.Bands(g).Columns("CONTENTS").ValueList = vListCompSubstance
                    End If

                    '' Fuel Type
                    'If ugTanks.DisplayLayout.Bands(g).Columns("FUEL TYPE").ValueList Is Nothing Then
                    '    vListCompFuelType = New Infragistics.Win.ValueList
                    '    For Each row As DataRow In oTank.PopulateCompartmentFuelTypes.Rows
                    '        vListCompFuelType.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                    '    Next
                    '    ugTanks.DisplayLayout.Bands(g).Columns("FUEL TYPE").ValueList = vListCompFuelType
                    'End If

                    ' Material Of Construction
                    If ugTanks.DisplayLayout.Bands(g).Columns("MATERIALS").ValueList Is Nothing Then
                        vListMaterialTank = New Infragistics.Win.ValueList
                        For Each row As DataRow In oTank.PopulateTankMaterialOfConstruction.Rows
                            vListMaterialTank.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugTanks.DisplayLayout.Bands(g).Columns("MATERIALS").ValueList = vListMaterialTank
                    End If

                    ' Over Fill
                    If ugTanks.DisplayLayout.Bands(g).Columns("OVERFILL TYPE").ValueList Is Nothing Then
                        vListOverFill = New Infragistics.Win.ValueList
                        For Each row As DataRow In oTank.PopulateTankOverFillProtectionType.Rows
                            vListOverFill.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugTanks.DisplayLayout.Bands(g).Columns("OVERFILL TYPE").ValueList = vListOverFill
                    End If
                Next g


                For Each dr In ugTanks.Rows
                    SetupTankPipeTermRow(TankPipeTermGrid.Tank, dr, [readOnly], oTank)

                Next

                'If CPTypeCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("CP TYPE").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("CP TYPE").Hidden = False
                'End If
                'If CPInstalledCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("CP INSTALLED").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("CP INSTALLED").Hidden = False
                'End If
                'If CPTestedCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("CP TESTED").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("CP TESTED").Hidden = False
                'End If
                'If LeakDetectionCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("LEAK DETECTION").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("LEAK DETECTION").Hidden = False
                'End If
                'If OverFillCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("OVERFILL TYPE").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("OVERFILL TYPE").Hidden = False
                'End If
                'If LinedCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("LINED").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("LINED").Hidden = False
                'End If
                'If LiningInspectedCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("LINING INSPECTED").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("LINING INSPECTED").Hidden = False
                'End If
                'If PTTCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("PTT").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("PTT").Hidden = False
                'End If
                'If LastUsedCount = ugTanks.Rows.Count Then
                '    ugTanks.DisplayLayout.Bands(0).Columns("LAST USED").Hidden = True
                'Else
                '    ugTanks.DisplayLayout.Bands(0).Columns("LAST USED").Hidden = False
                'End If
            End If

            If grid = TankPipeTermGrid.Pipe Or grid = TankPipeTermGrid.All Then
                '************************************************************************
                ' Pipe
                '************************************************************************

                If oTank Is Nothing Then
                    oTank = New MUSTER.BusinessLogic.pTank
                End If
                If oPipe Is Nothing Then
                    oPipe = New MUSTER.BusinessLogic.pPipe
                End If

                'InitSetupTankPipeTermVariables()

                ugPipes.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free

                ugPipes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
                ugPipes.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No

                For g As Integer = 0 To 1
                    ugPipes.DisplayLayout.Bands(g).Columns("PARENT_PIPE_ID").Hidden = True
                    ugPipes.DisplayLayout.Bands(g).Columns("Pipe #").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                    ugPipes.DisplayLayout.Bands(g).Columns("FACILITY_ID").Hidden = True
                    ugPipes.DisplayLayout.Bands(g).Columns("TANK_ID").Hidden = True
                    ugPipes.DisplayLayout.Bands(g).Columns("COMPARTMENT_NUMBER").Hidden = True
                    ugPipes.DisplayLayout.Bands(g).Columns("PIPE_ID").Hidden = True
                    ugPipes.DisplayLayout.Bands(g).Columns("POSITION").Hidden = True

                    ' specifying which cols are drop down
                    ugPipes.DisplayLayout.Bands(g).Columns("STATUS").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugPipes.DisplayLayout.Bands(g).Columns("TYPE OF SYSTEM").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugPipes.DisplayLayout.Bands(g).Columns("MATERIALS").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugPipes.DisplayLayout.Bands(g).Columns("OPTIONS").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugPipes.DisplayLayout.Bands(g).Columns("BRAND").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugPipes.DisplayLayout.Bands(g).Columns("CP TYPE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugPipes.DisplayLayout.Bands(g).Columns("PRI. LEAK DETECTION").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugPipes.DisplayLayout.Bands(g).Columns("SEC. LEAK DETECTION").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

                Next g
                ' populate columns
                If ugPipes.DisplayLayout.ValueLists.All.Length = 0 Then
                    ugPipes.DisplayLayout.ValueLists.Add("PIPE_STATUS")
                    ugPipes.DisplayLayout.ValueLists.Add("TYPE_OF_SYSTEM")
                    ugPipes.DisplayLayout.ValueLists.Add("MATERIAL_OF_CONSTRUCTION")
                    ugPipes.DisplayLayout.ValueLists.Add("CONSTRUCTION_OPTIONS")
                    ugPipes.DisplayLayout.ValueLists.Add("BRAND_OF_PIPE")
                    ugPipes.DisplayLayout.ValueLists.Add("CP_TYPE")
                    ugPipes.DisplayLayout.ValueLists.Add("PRIMARY_LEAK_DET")
                    ugPipes.DisplayLayout.ValueLists.Add("SECONDARY_LEAK_DET")
                End If

                For g As Integer = 0 To 1
                    ' populate the whole column as the table is the same for each row
                    ' Type of Pipe System
                    If ugPipes.DisplayLayout.Bands(g).Columns("TYPE OF SYSTEM").ValueList Is Nothing Then
                        vListType = New Infragistics.Win.ValueList
                        For Each row As DataRow In oPipe.PopulatePipeType.Rows
                            vListType.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugPipes.DisplayLayout.Bands(g).Columns("TYPE OF SYSTEM").ValueList = vListType
                    End If

                    ' Material Of Construction
                    If ugPipes.DisplayLayout.Bands(g).Columns("MATERIALS").ValueList Is Nothing Then
                        vListMaterialPipe = New Infragistics.Win.ValueList
                        For Each row As DataRow In oPipe.PopulatePipeMaterialOfConstruction.Rows
                            vListMaterialPipe.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugPipes.DisplayLayout.Bands(g).Columns("MATERIALS").ValueList = vListMaterialPipe
                    End If

                    ' Brand
                    If ugPipes.DisplayLayout.Bands(g).Columns("BRAND").ValueList Is Nothing Then
                        vListBrand = New Infragistics.Win.ValueList
                        For Each row As DataRow In oPipe.PopulatePipeManufacturer.Rows
                            vListBrand.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugPipes.DisplayLayout.Bands(g).Columns("BRAND").ValueList = vListBrand
                    End If

                    ' CP Type
                    If ugPipes.DisplayLayout.Bands(g).Columns("CP TYPE").ValueList Is Nothing Then
                        vListCPType = New Infragistics.Win.ValueList
                        For Each row As DataRow In oPipe.PopulatePipeCPType.Rows
                            vListCPType.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugPipes.DisplayLayout.Bands(g).Columns("CP TYPE").ValueList = vListCPType
                    End If

                Next g

                For Each dr In ugPipes.Rows
                    SetupTankPipeTermRow(TankPipeTermGrid.Pipe, dr, [readOnly], oTank, oPipe)

                    If dr.HasChild Then
                        dr.Appearance.BackColor = Color.LightGoldenrodYellow

                        For Each dr2 As Infragistics.Win.UltraWinGrid.UltraGridRow In dr.ChildBands(0).Rows
                            dr2.Appearance.BackColor = Color.LightSkyBlue
                        Next
                    End If
                Next

                'If BrandCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("BRAND").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("BRAND").Hidden = False
                'End If
                'If CPTypeCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("CP TYPE").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("CP TYPE").Hidden = False
                'End If
                'If CPInstalledCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("CP INSTALLED").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("CP INSTALLED").Hidden = False
                'End If
                'If CPTestedCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("CP TESTED").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("CP TESTED").Hidden = False
                'End If
                'If PriLeakDetectionCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("PRI. LEAK DETECTION").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("PRI. LEAK DETECTION").Hidden = False
                'End If
                'If SecLeakDetectionCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("SEC. LEAK DETECTION").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("SEC. LEAK DETECTION").Hidden = False
                'End If
                'If ALLDTestedCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("ALLD TESTED").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("ALLD TESTED").Hidden = False
                'End If
                'If PTTCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("PTT").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("PTT").Hidden = False
                'End If
                'If PipeLastUsedCount = ugPipes.Rows.Count Then
                '    ugPipes.DisplayLayout.Bands(0).Columns("LAST USED").Hidden = True
                'Else
                '    ugPipes.DisplayLayout.Bands(0).Columns("LAST USED").Hidden = False
                'End If
            End If

            If grid = TankPipeTermGrid.Term Or grid = TankPipeTermGrid.All Then
                '************************************************************************
                ' Term
                '************************************************************************

                If oTank Is Nothing Then
                    oTank = New MUSTER.BusinessLogic.pTank
                End If
                If oPipe Is Nothing Then
                    oPipe = New MUSTER.BusinessLogic.pPipe
                End If

                'InitSetupTankPipeTermVariables()

                ugTerminations.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free

                ugTerminations.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
                ugTerminations.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No

                For g As Integer = 0 To 1

                    ugTerminations.DisplayLayout.Bands(g).Columns("Pipe #").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ugTerminations.DisplayLayout.Bands(g).Columns("PARENT_PIPE_ID").Hidden = True

                    ugTerminations.DisplayLayout.Bands(g).Columns("FACILITY_ID").Hidden = True
                    ugTerminations.DisplayLayout.Bands(g).Columns("TANK_ID").Hidden = True
                    ugTerminations.DisplayLayout.Bands(g).Columns("COMPARTMENT_NUMBER").Hidden = True
                    ugTerminations.DisplayLayout.Bands(g).Columns("PIPE_ID").Hidden = True
                    ugTerminations.DisplayLayout.Bands(g).Columns("POSITION").Hidden = True

                    ' specifying which cols are drop down
                    ugTerminations.DisplayLayout.Bands(g).Columns("TYPE TERM@TANK").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTerminations.DisplayLayout.Bands(g).Columns("TYPE TERM@DISP").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTerminations.DisplayLayout.Bands(g).Columns("TANK TERM CP").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    ugTerminations.DisplayLayout.Bands(g).Columns("DISP. TERM CP").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

                Next g
                ' populate columns
                If ugTerminations.DisplayLayout.ValueLists.All.Length = 0 Then
                    ugTerminations.DisplayLayout.ValueLists.Add("TYPE_TERM_TANK")
                    ugTerminations.DisplayLayout.ValueLists.Add("TYPE_TERM_DISP")
                    ugTerminations.DisplayLayout.ValueLists.Add("TANK_TERM_CPTYPE")
                    ugTerminations.DisplayLayout.ValueLists.Add("DISP_TERM_CPTYPE")
                End If

                For g As Integer = 0 To 1
                    ' populate the whole column as the table is the same for each row
                    ' Type Term @ Dispenser
                    If ugTerminations.DisplayLayout.Bands(g).Columns("TYPE TERM@DISP").ValueList Is Nothing Then
                        vListTypeTermDisp = New Infragistics.Win.ValueList
                        For Each row As DataRow In oPipe.PopulatePipeTerminationDispenserType.Rows
                            vListTypeTermDisp.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugTerminations.DisplayLayout.Bands(g).Columns("TYPE TERM@DISP").ValueList = vListTypeTermDisp
                    End If

                    ' Type Term @ Tank
                    If ugTerminations.DisplayLayout.Bands(g).Columns("TYPE TERM@TANK").ValueList Is Nothing Then
                        vListTypeTermTank = New Infragistics.Win.ValueList
                        For Each row As DataRow In oPipe.PopulatePipeTerminationTankType.Rows
                            vListTypeTermTank.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugTerminations.DisplayLayout.Bands(g).Columns("TYPE TERM@TANK").ValueList = vListTypeTermTank
                    End If

                    ' Dispenser Term CP
                    If ugTerminations.DisplayLayout.Bands(g).Columns("DISP. TERM CP").ValueList Is Nothing Then
                        vListTDispTermCP = New Infragistics.Win.ValueList
                        For Each row As DataRow In oPipe.PopulatePipeTerminationDispenserCPType.Rows
                            vListTDispTermCP.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugTerminations.DisplayLayout.Bands(g).Columns("DISP. TERM CP").ValueList = vListTDispTermCP
                    End If

                    ' Tank Term CP
                    If ugTerminations.DisplayLayout.Bands(g).Columns("TANK TERM CP").ValueList Is Nothing Then
                        vListTankTermCP = New Infragistics.Win.ValueList
                        For Each row As DataRow In oPipe.PopulatePipeTerminationTankCPType.Rows
                            vListTankTermCP.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        ugTerminations.DisplayLayout.Bands(g).Columns("TANK TERM CP").ValueList = vListTankTermCP
                    End If

                Next g

                For Each dr In ugTerminations.Rows
                    SetupTankPipeTermRow(TankPipeTermGrid.Term, dr, [readOnly], oTank, oPipe)

                    If dr.HasChild Then
                        dr.Appearance.BackColor = Color.LightGoldenrodYellow

                        For Each dr2 As Infragistics.Win.UltraWinGrid.UltraGridRow In dr.ChildBands(0).Rows
                            dr2.Appearance.BackColor = Color.LightSkyBlue
                        Next
                    End If
                Next

                'If TermCPTestedCount = ugTerminations.Rows.Count Then
                '    ugTerminations.DisplayLayout.Bands(0).Columns("TERM CP TESTED").Hidden = True
                'Else
                '    ugTerminations.DisplayLayout.Bands(0).Columns("TERM CP TESTED").Hidden = False
                'End If
                'If TermDispCPTypeCount = ugTerminations.Rows.Count Then
                '    ugTerminations.DisplayLayout.Bands(0).Columns("DISP. TERM CP").Hidden = True
                'Else
                '    ugTerminations.DisplayLayout.Bands(0).Columns("DISP. TERM CP").Hidden = False
                'End If
                'If TermTankCPTypeCount = ugTerminations.Rows.Count Then
                '    ugTerminations.DisplayLayout.Bands(0).Columns("TANK TERM CP").Hidden = True
                'Else
                '    ugTerminations.DisplayLayout.Bands(0).Columns("TANK TERM CP").Hidden = False
                'End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SetupTankPipeTermRow(ByVal grid As TankPipeTermGrid, ByRef dr As Infragistics.Win.UltraWinGrid.UltraGridRow, Optional ByVal [readOnly] As Boolean = False, Optional ByVal pTank As MUSTER.BusinessLogic.pTank = Nothing, Optional ByVal pPipe As MUSTER.BusinessLogic.pPipe = Nothing)
        Dim dtOfInspection As Date
        Dim pipeID As String
        'Dim bolIsFacCapCandidate As Boolean = oInspection.CheckListMaster.Owner.Facilities.CAPCandidate
        Dim bolUpdateObject As Boolean = Not [readOnly]
        Try
            If Date.Compare(oInspection.RescheduledDate, dtNullDate) = 0 Then
                dtOfInspection = oInspection.ScheduledDate
            Else
                dtOfInspection = oInspection.RescheduledDate
            End If
            If grid = TankPipeTermGrid.Tank Then
                oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(CType(dr.Cells("TANK_ID").Value, Integer))
                pTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If Not (dr.Cells("COMPARTMENT_NUMBER").Value Is DBNull.Value) Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, pTank.TankId.ToString + "|" + dr.Cells("COMPARTMENT_NUMBER").Value.ToString)
                End If
                If pTank.TankStatus = 426 Then ' POU
                    bolUpdateObject = True
                End If
                ' Enable / Disable fields acc to Registration Use Case

                ' Contents
                If vListCompSubstance.FindByDataValue(dr.Cells("CONTENTS").Value) Is Nothing Then
                    dr.Cells("CONTENTS").Value = DBNull.Value
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.Compartments.Substance <> 0 Then
                            pTank.Compartments.Substance = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("CONTENTS").Value Is DBNull.Value Then
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        pTank.Compartments.Substance = 0
                        regenerateCheckListItems = True
                    End If
                ElseIf dr.Cells("CONTENTS").Value <> pTank.Compartments.Substance Then
                    If pTank.Compartments.Substance <= 0 Then
                        dr.Cells("CONTENTS").Value = DBNull.Value
                    Else
                        dr.Cells("CONTENTS").Value = pTank.Compartments.Substance
                    End If
                End If

                ' Fuel Type
                If dr.Cells("CONTENTS").Value Is DBNull.Value Then
                    dr.Cells("FUEL TYPE").Value = DBNull.Value
                    dr.Cells("FUEL TYPE").Hidden = True
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.Compartments.FuelTypeId <> 0 Then
                            pTank.Compartments.FuelTypeId = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                Else
                    Dim dtFuelType As DataTable = pTank.PopulateCompartmentFuelTypes(dr.Cells("CONTENTS").Value)
                    If dtFuelType Is Nothing Then
                        dr.Cells("FUEL TYPE").Value = DBNull.Value
                        dr.Cells("FUEL TYPE").Hidden = True
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pTank.Compartments.FuelTypeId <> 0 Then
                                pTank.Compartments.FuelTypeId = 0
                                regenerateCheckListItems = True
                            End If
                        End If
                    Else
                        dr.Cells("FUEL TYPE").Hidden = False
                        Dim vListCompFuelType As New Infragistics.Win.ValueList
                        For Each row As DataRow In dtFuelType.Rows
                            vListCompFuelType.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        dr.Cells("FUEL TYPE").ValueList = vListCompFuelType

                        If vListCompFuelType.FindByDataValue(dr.Cells("FUEL TYPE").Value) Is Nothing Then
                            dr.Cells("FUEL TYPE").Value = DBNull.Value
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If pTank.Compartments.FuelTypeId <> 0 Then
                                    pTank.Compartments.FuelTypeId = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        ElseIf dr.Cells("FUEL TYPE").Value Is DBNull.Value Then
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                pTank.Compartments.FuelTypeId = 0
                                regenerateCheckListItems = True
                            End If
                        ElseIf dr.Cells("FUEL TYPE").Value <> pTank.Compartments.FuelTypeId Then
                            If pTank.Compartments.FuelTypeId <= 0 Then
                                dr.Cells("FUEL TYPE").Value = DBNull.Value
                            Else
                                dr.Cells("FUEL TYPE").Value = pTank.Compartments.FuelTypeId
                            End If
                        End If

                    End If
                End If

                ' Material Of Construction
                If vListMaterialTank.FindByDataValue(dr.Cells("MATERIALS").Value) Is Nothing Then
                    dr.Cells("MATERIALS").Value = DBNull.Value
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.TankMatDesc <> 0 Then
                            pTank.TankMatDesc = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("MATERIALS").Value Is DBNull.Value Then
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        pTank.TankMatDesc = 0
                        regenerateCheckListItems = True
                    End If
                ElseIf dr.Cells("MATERIALS").Value <> pTank.TankMatDesc Then
                    If pTank.TankMatDesc <= 0 Then
                        dr.Cells("MATERIALS").Value = DBNull.Value
                    Else
                        dr.Cells("MATERIALS").Value = pTank.TankMatDesc
                    End If
                End If

                ' Construction Options
                Dim dt As New DataTable
                Dim vListOptions As New Infragistics.Win.ValueList
                dt = pTank.PopulateTankSecondaryOption(IIf(dr.Cells("MATERIALS").Value Is DBNull.Value, 0, dr.Cells("MATERIALS").Value))
                If Not (dt Is Nothing) Then
                    dr.Cells("OPTIONS").Hidden = False
                    For Each row As DataRow In dt.Rows
                        vListOptions.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                    Next
                End If
                dr.Cells("OPTIONS").ValueList = vListOptions
                If vListOptions.FindByDataValue(dr.Cells("OPTIONS").Value) Is Nothing Then
                    dr.Cells("OPTIONS").Value = DBNull.Value
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.TankModDesc <> 0 Then
                            pTank.TankModDesc = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("OPTIONS").Value Is DBNull.Value Then
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        pTank.TankModDesc = 0
                        regenerateCheckListItems = True
                    End If
                ElseIf dr.Cells("OPTIONS").Value <> pTank.TankModDesc Then
                    If pTank.TankModDesc <= 0 Then
                        dr.Cells("OPTIONS").Value = DBNull.Value
                    Else
                        dr.Cells("OPTIONS").Value = pTank.TankModDesc
                    End If
                End If

                ' Enable Field              Condition
                ' Tank CP Type              Tank Mod Desc (Tank Sec Option) like 'Cathodically Protected'
                ' Tank CP Installed         Tank Mod Desc (Tank Sec Option) like 'Cathodically Protected'
                ' Tank CP Last Tested       Tank Mod Desc (Tank Sec Option) like 'Cathodically Protected'
                ' Lined Interior Install    Tank Mod Desc (Tank Sec Option) like 'Lined'
                ' Lined Interior Inspect    Tank Mod Desc (Tank Sec Option) like 'Lined'
                If Not vListOptions.FindByDataValue(dr.Cells("OPTIONS").Value) Is Nothing Then
                    If vListOptions.FindByDataValue(dr.Cells("OPTIONS").Value).DisplayText.IndexOf("Cathodically Protected") > -1 Then
                        dr.Cells("CP TYPE").Hidden = False
                        dr.Cells("CP INSTALLED").Hidden = False
                        dr.Cells("CP TESTED").Hidden = False
                        ' per bug 2573, need to clear date even if fac is not cap candidate
                        If Not oInspection.CAPDatesEntered Then ' If bolIsFacCapCandidate And Not oInspection.CAPDatesEntered Then
                            dr.Cells("CP TESTED").Value = DBNull.Value
                            'Manju
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If Date.Compare(pTank.LastTCPDate, dtNullDate) <> 0 Then
                                    pTank.LastTCPDate = dtNullDate
                                    regenerateCheckListItems = True
                                End If
                                'Else
                                '    pTank.LastTCPDate = dtNullDate
                            End If
                        End If

                        'LinedCount += 1
                        'LiningInspectedCount += 1
                        dr.Cells("LINED").Hidden = True
                        dr.Cells("LINING INSPECTED").Hidden = True

                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If Date.Compare(pTank.LinedInteriorInstallDate, dtNullDate) <> 0 Or _
                                Date.Compare(pTank.LinedInteriorInspectDate, dtNullDate) <> 0 Then
                                pTank.LinedInteriorInstallDate = dtNullDate
                                pTank.LinedInteriorInspectDate = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If

                        ' CPType
                        Dim vListCPType As New Infragistics.Win.ValueList
                        For Each row As DataRow In pTank.PopulateTankCPType.Rows
                            vListCPType.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        dr.Cells("CP TYPE").ValueList = vListCPType
                        If vListCPType.FindByDataValue(dr.Cells("CP TYPE").Value) Is Nothing Then
                            dr.Cells("CP TYPE").Value = DBNull.Value
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If pTank.TankCPType <> 0 Then
                                    pTank.TankCPType = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        ElseIf dr.Cells("CP TYPE").Value Is DBNull.Value Then
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                pTank.TankCPType = 0
                                regenerateCheckListItems = True
                            End If
                        ElseIf dr.Cells("CP TYPE").Value <> pTank.TankCPType Then
                            If pTank.TankCPType <= 0 Then
                                dr.Cells("CP TYPE").Value = DBNull.Value
                            Else
                                dr.Cells("CP TYPE").Value = pTank.TankCPType
                            End If
                        End If
                    ElseIf vListOptions.FindByDataValue(dr.Cells("OPTIONS").Value).DisplayText.IndexOf("Lined") > -1 Then
                        dr.Cells("LINED").Hidden = False
                        dr.Cells("LINING INSPECTED").Hidden = False

                        If Not oInspection.CAPDatesEntered Then
                            dr.Cells("LINING INSPECTED").Value = DBNull.Value
                            'Manju
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If Date.Compare(pTank.LinedInteriorInspectDate, dtNullDate) <> 0 Then
                                    pTank.LinedInteriorInspectDate = dtNullDate
                                    regenerateCheckListItems = True
                                End If
                                'Else
                                '    pTank.LinedInteriorInspectDate = dtNullDate
                            End If
                        End If

                        'CPTypeCount += 1
                        'CPInstalledCount += 1
                        'CPTestedCount += 1
                        dr.Cells("CP TYPE").Hidden = True
                        dr.Cells("CP INSTALLED").Hidden = True
                        dr.Cells("CP TESTED").Hidden = True

                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pTank.TankCPType <> 0 Or _
                                Date.Compare(pTank.TCPInstallDate, dtNullDate) <> 0 Or _
                                Date.Compare(pTank.LastTCPDate, dtNullDate) <> 0 Then
                                pTank.TankCPType = 0
                                pTank.TCPInstallDate = dtNullDate
                                pTank.LastTCPDate = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                    Else
                        'CPTypeCount += 1
                        'CPInstalledCount += 1
                        'CPTestedCount += 1
                        'LinedCount += 1
                        'LiningInspectedCount += 1
                        dr.Cells("CP TYPE").Hidden = True
                        dr.Cells("CP INSTALLED").Hidden = True
                        dr.Cells("CP TESTED").Hidden = True
                        dr.Cells("LINED").Hidden = True
                        dr.Cells("LINING INSPECTED").Hidden = True

                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pTank.TankCPType <> 0 Or _
                                Date.Compare(pTank.TCPInstallDate, dtNullDate) <> 0 Or _
                                Date.Compare(pTank.LastTCPDate, dtNullDate) <> 0 Or _
                                Date.Compare(pTank.LinedInteriorInstallDate, dtNullDate) <> 0 Or _
                                Date.Compare(pTank.LinedInteriorInspectDate, dtNullDate) <> 0 Then
                                pTank.TankCPType = 0
                                pTank.TCPInstallDate = dtNullDate
                                pTank.LastTCPDate = dtNullDate
                                pTank.LinedInteriorInstallDate = dtNullDate
                                pTank.LinedInteriorInspectDate = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                    End If
                Else
                    'CPTypeCount += 1
                    'CPInstalledCount += 1
                    'CPTestedCount += 1
                    'LinedCount += 1
                    'LiningInspectedCount += 1
                    dr.Cells("CP TYPE").Hidden = True
                    dr.Cells("CP INSTALLED").Hidden = True
                    dr.Cells("CP TESTED").Hidden = True
                    dr.Cells("LINED").Hidden = True
                    dr.Cells("LINING INSPECTED").Hidden = True

                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.TankCPType <> 0 Or _
                            Date.Compare(pTank.TCPInstallDate, dtNullDate) <> 0 Or _
                            Date.Compare(pTank.LastTCPDate, dtNullDate) <> 0 Or _
                            Date.Compare(pTank.LinedInteriorInstallDate, dtNullDate) <> 0 Or _
                            Date.Compare(pTank.LinedInteriorInspectDate, dtNullDate) <> 0 Then
                            pTank.TankCPType = 0
                            pTank.TCPInstallDate = dtNullDate
                            pTank.LastTCPDate = dtNullDate
                            pTank.LinedInteriorInstallDate = dtNullDate
                            pTank.LinedInteriorInspectDate = dtNullDate
                            regenerateCheckListItems = True
                        End If
                    End If
                End If

                ' Need to enable Release Detection (Tank LD) Only if there is some value in Tank Mod Desc
                ' cause it returns no values if = 0
                Dim bolHideLeakDetection As Boolean = False
                If dr.Cells("OPTIONS").Value Is DBNull.Value Then
                    If dr.Tag Is Nothing Then
                        bolHideLeakDetection = True
                    ElseIf dr.Tag = 0 Then
                        bolHideLeakDetection = True
                    End If
                    'dr.Cells("LEAK DETECTION").Hidden = True
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.TankLD <> 0 Then
                            pTank.TankLD = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                    'LeakDetectionCount += 1
                ElseIf dr.Cells("OPTIONS").Value = 0 Then
                    If dr.Tag Is Nothing Then
                        bolHideLeakDetection = True
                    ElseIf dr.Tag = 0 Then
                        bolHideLeakDetection = True
                    End If
                    'dr.Cells("LEAK DETECTION").Hidden = True
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.TankLD <> 0 Then
                            pTank.TankLD = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                    'LeakDetectionCount += 1
                Else
                    dr.Cells("LEAK DETECTION").Hidden = False
                    ' Leak Detection
                    Dim vListLeakDetection As New Infragistics.Win.ValueList
                    dt = pTank.PopulateTankReleaseDetection(dr.Cells("OPTIONS").Value, IIf(dr.Cells("SIZE").Value Is DBNull.Value, 0, dr.Cells("SIZE").Value))
                    If Not (dt Is Nothing) Then
                        For Each row As DataRow In dt.Rows
                            vListLeakDetection.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                    End If
                    dr.Cells("LEAK DETECTION").ValueList = vListLeakDetection
                    If vListLeakDetection.FindByDataValue(dr.Cells("LEAK DETECTION").Value) Is Nothing Then
                        dr.Cells("LEAK DETECTION").Value = DBNull.Value
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pTank.TankLD <> 0 Then
                                pTank.TankLD = 0
                                regenerateCheckListItems = True
                            End If
                        End If
                    ElseIf dr.Cells("LEAK DETECTION").Value Is DBNull.Value Then
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            pTank.TankLD = 0
                            regenerateCheckListItems = True
                        End If
                    ElseIf dr.Cells("LEAK DETECTION").Value <> pTank.TankLD Then
                        If pTank.TankLD <= 0 Then
                            dr.Cells("LEAK DETECTION").Value = DBNull.Value
                        Else
                            dr.Cells("LEAK DETECTION").Value = pTank.TankLD
                        End If
                    End If

                    ' Enable Field              Condition
                    ' TTT Date                  Tank LD = Inventory Control / Precision Tightness Test
                    ' Droptube for IC           Tank LD = Inventory Control / Precision Tightness Test (Automatically check true)
                    If Not vListLeakDetection.FindByDataValue(dr.Cells("LEAK DETECTION").Value) Is Nothing Then
                        If vListLeakDetection.FindByDataValue(dr.Cells("LEAK DETECTION").Value).DisplayText.IndexOf("Inventory Control/Precision Tightness Testing") > -1 Then
                            dr.Cells("PTT").Hidden = False
                            If Not oInspection.CAPDatesEntered Then
                                dr.Cells("PTT").Value = DBNull.Value
                                'Manju
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    If Date.Compare(pTank.TTTDate, dtNullDate) <> 0 Then
                                        pTank.TTTDate = dtNullDate
                                        regenerateCheckListItems = True
                                    End If
                                    'Else
                                    '    pTank.TTTDate = dtNullDate
                                End If
                            End If
                        Else
                            dr.Cells("PTT").Hidden = True
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If Date.Compare(pTank.TTTDate, dtNullDate) <> 0 Then
                                    pTank.TTTDate = dtNullDate
                                    regenerateCheckListItems = True
                                End If
                            End If
                            'PTTCount += 1
                        End If
                    Else
                        dr.Cells("PTT").Hidden = True
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If Date.Compare(pTank.TTTDate, dtNullDate) <> 0 Then
                                pTank.TTTDate = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                        'PTTCount += 1
                    End If
                End If

                If bolHideLeakDetection Then
                    dr.Cells("LEAK DETECTION").Hidden = True
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.TankLD <> 0 Then
                            pTank.TankLD = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                End If

                If pTank.SmallDelivery Then
                    'OverFillCount += 1
                    dr.Cells("OVERFILL TYPE").Hidden = True
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pTank.OverFillType <> 0 Then
                            pTank.OverFillType = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                Else
                    ' #2889
                    If Not pTank.Compartment And pTank.Compartments.Substance = 314 Then
                        dr.Cells("OVERFILL TYPE").Hidden = True
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pTank.OverFillType <> 0 Then
                                pTank.OverFillType = 0
                                regenerateCheckListItems = True
                            End If
                        End If
                    Else
                        dr.Cells("OVERFILL TYPE").Hidden = False
                        ' Over Fill
                        If vListOverFill.FindByDataValue(dr.Cells("OVERFILL TYPE").Value) Is Nothing Then
                            dr.Cells("OVERFILL TYPE").Value = DBNull.Value
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If pTank.OverFillType <> 0 Then
                                    pTank.OverFillType = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        End If
                    End If
                End If

                ' if ciu disable date LAST USED
                If dr.Cells("STATUS").Text.IndexOf("Currently In Use") > -1 Then
                    dr.Cells("LAST USED").Hidden = True
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If Date.Compare(pTank.DateLastUsed, dtNullDate) <> 0 Then
                            pTank.DateLastUsed = dtNullDate
                            regenerateCheckListItems = True
                        End If
                    End If
                    'LastUsedCount += 1
                Else
                    dr.Cells("LAST USED").Hidden = False
                End If

                ' make the row readonly and change the color if tank is pou
                If dr.Cells("STATUS").Text.IndexOf("Permanently Out of Use") > -1 Then
                    dr.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    dr.Appearance.BackColor = Color.LightGray
                End If
                ' make the row readonly if checklist is opened in readonly mode
                If [readOnly] Then
                    dr.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                End If
            ElseIf grid = TankPipeTermGrid.Pipe Then
                oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(CType(dr.Cells("TANK_ID").Value, Integer))
                pTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                'If pTank.Compartments.COMPARTMENTNumber <> dr.Cells("COMPARTMENT_NUMBER").Value Then
                pTank.Compartments.Retrieve(pTank.TankInfo, pTank.TankId.ToString + "|" + dr.Cells("COMPARTMENT_NUMBER").Value.ToString)
                'End If
                pipeID = dr.Cells("TANK_ID").Value.ToString + "|" + dr.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + dr.Cells("PIPE_ID").Value.ToString
                'oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.Pipes.Pipe = pTank.TankInfo.pipesCollection.Item(pipeID)
                oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.Pipes.Retrieve(pTank.TankInfo, pipeID)
                pPipe = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.Pipes

                ' Enable / Disable fields acc to Registration Use Case
                ' Status
                Dim dt As DataTable
                Dim vListStatus As New Infragistics.Win.ValueList
                For Each row As DataRow In pPipe.PopulatePipeStatus("EDIT").Rows
                    vListStatus.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                dr.Cells("STATUS").ValueList = vListStatus
                If vListStatus.FindByDataValue(dr.Cells("STATUS").Value) Is Nothing Then
                    dr.Cells("STATUS").Value = DBNull.Value
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pPipe.PipeStatusDesc <> 0 Then
                            pPipe.PipeStatusDesc = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("STATUS").Value Is DBNull.Value Then
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        pPipe.PipeStatusDesc = 0
                        regenerateCheckListItems = True
                    End If
                ElseIf dr.Cells("STATUS").Value <> pPipe.PipeStatusDesc Then
                    If pPipe.PipeStatusDesc <= 0 Then
                        dr.Cells("STATUS").Value = DBNull.Value
                    Else
                        dr.Cells("STATUS").Value = pPipe.PipeStatusDesc
                    End If
                End If

                ' Material
                If vListMaterialPipe.FindByDataValue(dr.Cells("MATERIALS").Value) Is Nothing Then
                    dr.Cells("MATERIALS").Value = DBNull.Value
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pPipe.PipeMatDesc <> 0 Then
                            pPipe.PipeMatDesc = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("MATERIALS").Value Is DBNull.Value Then
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        pPipe.PipeMatDesc = 0
                        regenerateCheckListItems = True
                    End If
                ElseIf dr.Cells("MATERIALS").Value <> pPipe.PipeMatDesc Then
                    If pPipe.PipeMatDesc <= 0 Then
                        dr.Cells("MATERIALS").Value = DBNull.Value
                    Else
                        dr.Cells("MATERIALS").Value = pPipe.PipeMatDesc
                    End If
                End If

                ' Construction Option
                Dim vListOptions As New Infragistics.Win.ValueList
                dt = New DataTable
                dt = pPipe.PopulatePipeSecondaryOptionNew(IIf(dr.Cells("STATUS").Value Is DBNull.Value, 0, dr.Cells("STATUS").Value), IIf(dr.Cells("MATERIALS").Value Is DBNull.Value, 0, dr.Cells("MATERIALS").Value))
                If Not (dt Is Nothing) Then
                    dr.Cells("OPTIONS").Hidden = False
                    For Each row As DataRow In dt.Rows
                        vListOptions.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                    Next
                End If
                dr.Cells("OPTIONS").ValueList = vListOptions
                If vListOptions.FindByDataValue(dr.Cells("OPTIONS").Value) Is Nothing Then
                    dr.Cells("OPTIONS").Value = DBNull.Value
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pPipe.PipeModDesc <> 0 Then
                            pPipe.PipeModDesc = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("OPTIONS").Value Is DBNull.Value Then
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        pPipe.PipeModDesc = 0
                        regenerateCheckListItems = True
                    End If
                ElseIf dr.Cells("OPTIONS").Value <> pPipe.PipeModDesc Then
                    If pPipe.PipeModDesc <= 0 Then
                        dr.Cells("OPTIONS").Value = DBNull.Value
                    Else
                        dr.Cells("OPTIONS").Value = pPipe.PipeModDesc
                    End If
                End If

                ' Enable Field              Condition
                ' Pipe CP Type              Pipe Mod Desc (Pipe Sec Option) like 'Cathodically Protected'
                ' Pipe CP Installed         Pipe Mod Desc (Pipe Sec Option) like 'Cathodically Protected'
                ' Pipe CP Last Tested       Pipe Mod Desc (Pipe Sec Option) like 'Cathodically Protected'
                If Not vListOptions.FindByDataValue(dr.Cells("OPTIONS").Value) Is Nothing Then
                    If vListOptions.FindByDataValue(dr.Cells("OPTIONS").Value).DisplayText.IndexOf("Cathodically Protected") > -1 Then
                        dr.Cells("CP TYPE").Hidden = False
                        dr.Cells("CP INSTALLED").Hidden = False
                        dr.Cells("CP TESTED").Hidden = False
                        If Not oInspection.CAPDatesEntered Then
                            dr.Cells("CP TESTED").Value = DBNull.Value
                            'Manju
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If Date.Compare(pPipe.PipeCPTest, dtNullDate) <> 0 Then
                                    pPipe.PipeCPTest = dtNullDate
                                    regenerateCheckListItems = True
                                End If
                                'Else
                                '    pPipe.PipeCPTest = dtNullDate
                            End If
                        End If
                        If vListCPType.FindByDataValue(dr.Cells("CP TYPE").Value) Is Nothing Then
                            dr.Cells("CP TYPE").Value = DBNull.Value
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If pPipe.PipeCPType <> 0 Then
                                    pPipe.PipeCPType = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        ElseIf dr.Cells("CP TYPE").Value Is DBNull.Value Then
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                pPipe.PipeCPType = 0
                                regenerateCheckListItems = True
                            End If
                        ElseIf dr.Cells("CP TYPE").Value <> pPipe.PipeCPType Then
                            If pPipe.PipeCPType <= 0 Then
                                dr.Cells("CP TYPE").Value = DBNull.Value
                            Else
                                dr.Cells("CP TYPE").Value = pPipe.PipeCPType
                            End If
                        End If
                    Else
                        dr.Cells("CP TYPE").Hidden = True
                        dr.Cells("CP INSTALLED").Hidden = True
                        dr.Cells("CP TESTED").Hidden = True
                        'CPTypeCount += 1
                        'CPInstalledCount += 1
                        'CPTestedCount += 1
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pPipe.PipeCPType <> 0 Or _
                                Date.Compare(pPipe.PipeCPInstalledDate, dtNullDate) <> 0 Or _
                                Date.Compare(pPipe.PipeCPTest, dtNullDate) <> 0 Then
                                pPipe.PipeCPType = 0
                                pPipe.PipeCPInstalledDate = dtNullDate
                                pPipe.PipeCPTest = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                    End If
                Else
                    dr.Cells("CP TYPE").Hidden = True
                    dr.Cells("CP INSTALLED").Hidden = True
                    dr.Cells("CP TESTED").Hidden = True
                    'CPTypeCount += 1
                    'CPInstalledCount += 1
                    'CPTestedCount += 1
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pPipe.PipeCPType <> 0 Or _
                            Date.Compare(pPipe.PipeCPInstalledDate, dtNullDate) <> 0 Or _
                            Date.Compare(pPipe.PipeCPTest, dtNullDate) <> 0 Then
                            pPipe.PipeCPType = 0
                            pPipe.PipeCPInstalledDate = dtNullDate
                            pPipe.PipeCPTest = dtNullDate
                            regenerateCheckListItems = True
                        End If
                    End If
                End If

                ' Manufacturer
                If vListBrand.FindByDataValue(dr.Cells("BRAND").Value) Is Nothing Then
                    dr.Cells("BRAND").Value = DBNull.Value
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pPipe.PipeManufacturer <> 0 Then
                            pPipe.PipeManufacturer = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("BRAND").Value Is DBNull.Value Then
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        pPipe.PipeManufacturer = 0
                        regenerateCheckListItems = True
                    End If
                ElseIf dr.Cells("BRAND").Value <> pPipe.PipeManufacturer Then
                    If pPipe.PipeManufacturer <= 0 Then
                        dr.Cells("BRAND").Value = DBNull.Value
                    Else
                        dr.Cells("BRAND").Value = pPipe.PipeManufacturer
                    End If
                End If

                ' Pipe Type
                If vListType.FindByDataValue(dr.Cells("TYPE OF SYSTEM").Value) Is Nothing Then
                    dr.Cells("TYPE OF SYSTEM").Value = DBNull.Value
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pPipe.PipeTypeDesc <> 0 Then
                            pPipe.PipeTypeDesc = 0
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("TYPE OF SYSTEM").Value Is DBNull.Value Then
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        pPipe.PipeTypeDesc = 0
                        regenerateCheckListItems = True
                    End If
                ElseIf dr.Cells("TYPE OF SYSTEM").Value <> pPipe.PipeTypeDesc Then
                    If pPipe.PipeTypeDesc <= 0 Then
                        dr.Cells("TYPE OF SYSTEM").Value = DBNull.Value
                    Else
                        dr.Cells("TYPE OF SYSTEM").Value = pPipe.PipeTypeDesc
                    End If
                End If

                ' PipeLD = Release Detection 1
                ' Alld Type = Release Detection 2

                ' Enable Field              Condition
                ' Release Detection 1       Pipe Type = 'Pressurized'
                ' Release Detection 1       Pipe Type = 'U.S. Suction'
                ' LTT Date                  Release Detection 1 = 'Line Tightness Testing'

                ' Release Detection 2       Pipe Type = 'Pressurized' and (PipeLD <> 'Continuous Interstitial Monitoring' OR 'Deferred')
                ' (ALLD Type)
                ' ALLD Test Date            Pipe Type = 'Pressurized' and (PipeLD <> 'Continuous Interstitial Monitoring' OR 'Deferred')
                '                                        and ALLD Type = 'Mechanical' and Pipe Status <> 'TOSI'
                If Not vListType.FindByDataValue(dr.Cells("TYPE OF SYSTEM").Value) Is Nothing Then
                    If vListType.FindByDataValue(dr.Cells("TYPE OF SYSTEM").Value).DisplayText.IndexOf("Pressurized") > -1 Then
                        dr.Cells("PRI. LEAK DETECTION").Hidden = False

                        'PopulatePipeReleaseDetection1
                        Dim vListPrimLeak As New Infragistics.Win.ValueList
                        dt = New DataTable
                        dt = pPipe.PopulatePipeReleaseDetection1(IIf(dr.Cells("OPTIONS").Value Is DBNull.Value, 0, dr.Cells("OPTIONS").Value))
                        If Not (dt Is Nothing) Then
                            For Each row As DataRow In dt.Rows
                                vListPrimLeak.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                            Next
                        End If
                        dr.Cells("PRI. LEAK DETECTION").ValueList = vListPrimLeak
                        If vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value) Is Nothing Then
                            dr.Cells("PRI. LEAK DETECTION").Value = DBNull.Value
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If pPipe.PipeLD <> 0 Then
                                    pPipe.PipeLD = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        ElseIf dr.Cells("PRI. LEAK DETECTION").Value Is DBNull.Value Then
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                pPipe.PipeLD = 0
                                regenerateCheckListItems = True
                            End If
                        ElseIf dr.Cells("PRI. LEAK DETECTION").Value <> pPipe.PipeLD Then
                            If pPipe.PipeLD <= 0 Then
                                dr.Cells("PRI. LEAK DETECTION").Value = DBNull.Value
                            Else
                                dr.Cells("PRI. LEAK DETECTION").Value = pPipe.PipeLD
                            End If
                        End If

                        If Not vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value) Is Nothing Then
                            If vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value).DisplayText.IndexOf("Line Tightness Testing") > -1 Then
                                dr.Cells("PTT").Hidden = False
                                If Not oInspection.CAPDatesEntered Then
                                    dr.Cells("PTT").Value = DBNull.Value
                                    'Manju
                                    If bolUpdateObject And Not bolCheckListPrinting Then
                                        If Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Then
                                            pPipe.LTTDate = dtNullDate
                                            regenerateCheckListItems = True
                                        End If
                                        'Else
                                        '    pPipe.LTTDate = dtNullDate
                                    End If
                                End If
                            ElseIf vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value).DisplayText.IndexOf("Electronic ALLD with 0.2") > -1 Then
                                dr.Cells("PTT").Hidden = True
                                'PTTCount += 1
                                'Electronic
                                dr.Cells("SEC. LEAK DETECTION").Value = 497
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    If Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Or _
                                        pPipe.ALLDType <> 497 Then
                                        pPipe.LTTDate = dtNullDate
                                        pPipe.ALLDType = 497
                                        regenerateCheckListItems = True
                                    End If
                                End If
                            ElseIf vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value).DisplayText.IndexOf("Continuous Interstitial Monitoring") > -1 Then
                                dr.Cells("PTT").Hidden = True
                                'PTTCount += 1
                                'Electronic
                                dr.Cells("SEC. LEAK DETECTION").Value = 498
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    If Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Or _
                                        pPipe.ALLDType <> 498 Then
                                        pPipe.LTTDate = dtNullDate
                                        pPipe.ALLDType = 498
                                        regenerateCheckListItems = True
                                    End If
                                End If
                            Else
                                dr.Cells("PTT").Hidden = True
                                'PTTCount += 1
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    If Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Then
                                        pPipe.LTTDate = dtNullDate
                                        regenerateCheckListItems = True
                                    End If
                                End If
                            End If

                            If Not vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value).DisplayText.IndexOf("Deferred") > -1 Then
                                dr.Cells("SEC. LEAK DETECTION").Hidden = False
                                'PopulatePipeReleaseDetection2
                                Dim vListSecLeak As New Infragistics.Win.ValueList
                                dt = New DataTable
                                dt = pPipe.PopulatePipeReleaseDetection2(IIf(dr.Cells("PRI. LEAK DETECTION").Value Is DBNull.Value, 0, dr.Cells("PRI. LEAK DETECTION").Value), _
                                                                            IIf(dr.Cells("OPTIONS").Value Is DBNull.Value, 0, dr.Cells("OPTIONS").Value))
                                If Not (dt Is Nothing) Then
                                    For Each row As DataRow In dt.Rows
                                        vListSecLeak.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                                    Next
                                End If
                                dr.Cells("SEC. LEAK DETECTION").ValueList = vListSecLeak
                                If vListSecLeak.FindByDataValue(dr.Cells("SEC. LEAK DETECTION").Value) Is Nothing Then
                                    dr.Cells("SEC. LEAK DETECTION").Value = DBNull.Value
                                    If bolUpdateObject And Not bolCheckListPrinting Then
                                        If pPipe.ALLDType <> 0 Then
                                            pPipe.ALLDType = 0
                                            regenerateCheckListItems = True
                                        End If
                                    End If
                                ElseIf dr.Cells("SEC. LEAK DETECTION").Value Is DBNull.Value Then
                                    If bolUpdateObject And Not bolCheckListPrinting Then
                                        pPipe.ALLDType = 0
                                        regenerateCheckListItems = True
                                    End If
                                ElseIf dr.Cells("SEC. LEAK DETECTION").Value <> pPipe.ALLDType Then
                                    If pPipe.ALLDType <= 0 Then
                                        dr.Cells("SEC. LEAK DETECTION").Value = DBNull.Value
                                    Else
                                        dr.Cells("SEC. LEAK DETECTION").Value = pPipe.ALLDType
                                    End If
                                End If

                                If Not vListSecLeak.FindByDataValue(dr.Cells("SEC. LEAK DETECTION").Value) Is Nothing Then
                                    If vListSecLeak.FindByDataValue(dr.Cells("SEC. LEAK DETECTION").Value).DisplayText.IndexOf("Mechanical") > -1 Then
                                        ' #2890 Disable ALLD Test Date if Pipe Status = TOSI
                                        If dr.Cells("STATUS").Text.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then
                                            dr.Cells("ALLD TESTED").Hidden = True
                                            dr.Cells("ALLD TESTED").Value = DBNull.Value
                                            If bolUpdateObject And Not bolCheckListPrinting Then
                                                If Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Then
                                                    pPipe.ALLDTestDate = dtNullDate
                                                    regenerateCheckListItems = True
                                                End If
                                            End If
                                        Else
                                            dr.Cells("ALLD TESTED").Hidden = False
                                            If Not oInspection.CAPDatesEntered Then
                                                dr.Cells("ALLD TESTED").Value = DBNull.Value
                                                'Manju
                                                If bolUpdateObject And Not bolCheckListPrinting Then
                                                    If Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Then
                                                        pPipe.ALLDTestDate = dtNullDate
                                                        regenerateCheckListItems = True
                                                    End If
                                                    'Else
                                                    '    pPipe.ALLDTestDate = dtNullDate
                                                End If
                                            End If
                                        End If
                                    ElseIf vListSecLeak.FindByDataValue(dr.Cells("SEC. LEAK DETECTION").Value).DisplayText.IndexOf("Electronic") > -1 Then
                                        dr.Cells("ALLD TESTED").Hidden = True
                                        'ALLDTestedCount += 1
                                        dr.Cells("PRI. LEAK DETECTION").Value = 246
                                        If bolUpdateObject And Not bolCheckListPrinting Then
                                            If Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Or _
                                                pPipe.PipeLD <> 246 Then
                                                pPipe.ALLDTestDate = dtNullDate
                                                pPipe.PipeLD = 246
                                                regenerateCheckListItems = True
                                            End If
                                        End If
                                    ElseIf vListSecLeak.FindByDataValue(dr.Cells("SEC. LEAK DETECTION").Value).DisplayText.IndexOf("Continuous") > -1 Then
                                        dr.Cells("ALLD TESTED").Hidden = True
                                        'ALLDTestedCount += 1
                                        dr.Cells("PRI. LEAK DETECTION").Value = 243
                                        If bolUpdateObject And Not bolCheckListPrinting Then
                                            If Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Or _
                                                pPipe.PipeLD <> 243 Then
                                                pPipe.ALLDTestDate = dtNullDate
                                                pPipe.PipeLD = 243
                                                regenerateCheckListItems = True
                                            End If
                                        End If
                                    Else
                                        dr.Cells("ALLD TESTED").Hidden = True
                                        'ALLDTestedCount += 1
                                        If bolUpdateObject And Not bolCheckListPrinting Then
                                            If Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Then
                                                pPipe.ALLDTestDate = dtNullDate
                                                regenerateCheckListItems = True
                                            End If
                                        End If
                                    End If
                                Else
                                    dr.Cells("ALLD TESTED").Hidden = True
                                    'ALLDTestedCount += 1
                                    If bolUpdateObject And Not bolCheckListPrinting Then
                                        If Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Then
                                            pPipe.ALLDTestDate = dtNullDate
                                            regenerateCheckListItems = True
                                        End If
                                    End If
                                End If
                            Else
                                dr.Cells("SEC. LEAK DETECTION").Hidden = True
                                'SecLeakDetectionCount += 1
                                dr.Cells("ALLD TESTED").Hidden = True
                                'ALLDTestedCount += 1
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    If pPipe.ALLDType <> 0 Or _
                                        Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Then
                                        pPipe.ALLDType = 0
                                        pPipe.ALLDTestDate = dtNullDate
                                        regenerateCheckListItems = True
                                    End If
                                End If
                            End If
                        Else
                            dr.Cells("PTT").Hidden = True
                            'PTTCount += 1
                            dr.Cells("SEC. LEAK DETECTION").Hidden = True
                            'SecLeakDetectionCount += 1
                            dr.Cells("ALLD TESTED").Hidden = True
                            'ALLDTestedCount += 1
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Or _
                                    pPipe.ALLDType <> 0 Or _
                                    Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Then
                                    pPipe.LTTDate = dtNullDate
                                    pPipe.ALLDType = 0
                                    pPipe.ALLDTestDate = dtNullDate
                                    regenerateCheckListItems = True
                                End If
                            End If
                        End If
                    ElseIf vListType.FindByDataValue(dr.Cells("TYPE OF SYSTEM").Value).DisplayText.IndexOf("U.S. Suction") > -1 Then
                        dr.Cells("PRI. LEAK DETECTION").Hidden = False

                        'PopulatePipeReleaseDetection1
                        Dim vListPrimLeak As New Infragistics.Win.ValueList
                        dt = New DataTable
                        dt = pPipe.PopulatePipeReleaseDetection1(IIf(dr.Cells("OPTIONS").Value Is DBNull.Value, 0, dr.Cells("OPTIONS").Value))
                        If Not (dt Is Nothing) Then
                            For Each row As DataRow In dt.Rows
                                vListPrimLeak.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                            Next
                        End If
                        dr.Cells("PRI. LEAK DETECTION").ValueList = vListPrimLeak
                        If vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value) Is Nothing Then
                            dr.Cells("PRI. LEAK DETECTION").Value = DBNull.Value
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If pPipe.PipeLD <> 0 Then
                                    pPipe.PipeLD = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        ElseIf dr.Cells("PRI. LEAK DETECTION").Value Is DBNull.Value Then
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                pPipe.PipeLD = 0
                                regenerateCheckListItems = True
                            End If
                        ElseIf dr.Cells("PRI. LEAK DETECTION").Value <> pPipe.PipeLD Then
                            If pPipe.PipeLD <= 0 Then
                                dr.Cells("PRI. LEAK DETECTION").Value = DBNull.Value
                            Else
                                dr.Cells("PRI. LEAK DETECTION").Value = pPipe.PipeLD
                            End If
                        End If

                        If Not vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value) Is Nothing Then
                            If vListPrimLeak.FindByDataValue(dr.Cells("PRI. LEAK DETECTION").Value).DisplayText.IndexOf("Line Tightness Testing") > -1 Then
                                dr.Cells("PTT").Hidden = False
                                If Not oInspection.CAPDatesEntered Then
                                    dr.Cells("PTT").Value = DBNull.Value
                                    'Manju
                                    If bolUpdateObject And Not bolCheckListPrinting Then
                                        If Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Then
                                            pPipe.LTTDate = dtNullDate
                                            regenerateCheckListItems = True
                                        End If
                                        'Else
                                        '    pPipe.LTTDate = dtNullDate
                                    End If
                                End If
                            Else
                                dr.Cells("PTT").Hidden = True
                                'PTTCount += 1
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    If Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Then
                                        pPipe.LTTDate = dtNullDate
                                        regenerateCheckListItems = True
                                    End If
                                End If
                            End If
                        Else
                            dr.Cells("PTT").Hidden = True
                            'PTTCount += 1
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Then
                                    pPipe.LTTDate = dtNullDate
                                    regenerateCheckListItems = True
                                End If
                            End If
                        End If

                        dr.Cells("SEC. LEAK DETECTION").Hidden = True
                        'SecLeakDetectionCount += 1
                        dr.Cells("ALLD TESTED").Hidden = True
                        'ALLDTestedCount += 1
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pPipe.ALLDType <> 0 Or _
                                Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Then
                                pPipe.ALLDType = 0
                                pPipe.ALLDTestDate = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                    Else
                        dr.Cells("PRI. LEAK DETECTION").Hidden = True
                        'PriLeakDetectionCount += 1
                        dr.Cells("SEC. LEAK DETECTION").Hidden = True
                        'SecLeakDetectionCount += 1
                        dr.Cells("PTT").Hidden = True
                        'PTTCount += 1
                        dr.Cells("ALLD TESTED").Hidden = True
                        'ALLDTestedCount += 1
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pPipe.PipeLD <> 0 Or _
                                pPipe.ALLDType <> 0 Or _
                                Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Or _
                                Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Then
                                pPipe.PipeLD = 0
                                pPipe.ALLDType = 0
                                pPipe.ALLDTestDate = dtNullDate
                                pPipe.LTTDate = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                    End If
                Else
                    dr.Cells("PRI. LEAK DETECTION").Hidden = True
                    'PriLeakDetectionCount += 1
                    dr.Cells("SEC. LEAK DETECTION").Hidden = True
                    'SecLeakDetectionCount += 1
                    dr.Cells("PTT").Hidden = True
                    'PTTCount += 1
                    dr.Cells("ALLD TESTED").Hidden = True
                    'ALLDTestedCount += 1
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If pPipe.PipeLD <> 0 Or _
                            pPipe.ALLDType <> 0 Or _
                            Date.Compare(pPipe.ALLDTestDate, dtNullDate) <> 0 Or _
                            Date.Compare(pPipe.LTTDate, dtNullDate) <> 0 Then
                            pPipe.PipeLD = 0
                            pPipe.ALLDType = 0
                            pPipe.ALLDTestDate = dtNullDate
                            pPipe.LTTDate = dtNullDate
                            regenerateCheckListItems = True
                        End If
                    End If
                End If

                If dr.Cells("STATUS").Text.IndexOf("Currently In Use") > -1 Then
                    dr.Cells("LAST USED").Hidden = True
                    'PipeLastUsedCount += 1
                    If bolUpdateObject And Not bolCheckListPrinting Then
                        If Date.Compare(pPipe.DateLastUsed, dtNullDate) <> 0 Then
                            pPipe.DateLastUsed = dtNullDate
                            regenerateCheckListItems = True
                        End If
                    End If
                ElseIf dr.Cells("STATUS").Text.IndexOf("Permanently Out of Use") > -1 Then
                    If pPipe.Pipe.POU And pPipe.Pipe.NonPre88 Then
                        dr.Cells("LAST USED").Hidden = True
                        'PipeLastUsedCount += 1
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If Date.Compare(pPipe.DateLastUsed, dtNullDate) <> 0 Then
                                pPipe.DateLastUsed = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                    End If
                    If Date.Compare(pPipe.DateLastUsed, CDate("12/22/1988")) >= 0 Then
                        dr.Cells("LAST USED").Hidden = True
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If Date.Compare(pPipe.DateLastUsed, dtNullDate) <> 0 Then
                                pPipe.DateLastUsed = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                    Else
                        dr.Cells("LAST USED").Hidden = False
                    End If
                Else
                    dr.Cells("LAST USED").Hidden = False
                    'PipeLastUsedCount += 1
                    'If bolUpdateObject And Not bolCheckListPrinting  Then
                    '    If Date.Compare(pPipe.DateLastUsed, dtNullDate) <> 0 Then
                    '        pPipe.DateLastUsed = dtNullDate
                    '        regenerateCheckListItems = True
                    '    End If
                    'End If
                End If

                ' make the row readonly if checklist is opened in readonly mode
                If [readOnly] Then
                    dr.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                End If
                ElseIf grid = TankPipeTermGrid.Term Then
                    oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(CType(dr.Cells("TANK_ID").Value, Integer))
                    pTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                    pipeID = dr.Cells("TANK_ID").Value.ToString + "|" + dr.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + dr.Cells("PIPE_ID").Value.ToString
                    oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.Pipes.Pipe = pTank.TankInfo.pipesCollection.Item(pipeID)
                    pPipe = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.Pipes

                    ' Enable / Disable fields acc to Registration Use Case
                    ' Type Term @ Dispenser
                    If vListTypeTermDisp.FindByDataValue(dr.Cells("TYPE TERM@DISP").Value) Is Nothing Then
                        dr.Cells("TYPE TERM@DISP").Value = DBNull.Value
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pPipe.TermTypeDisp <> 0 Then
                                pPipe.TermTypeDisp = 0
                                regenerateCheckListItems = True
                            End If
                        End If
                    ElseIf dr.Cells("TYPE TERM@DISP").Value Is DBNull.Value Then
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            pPipe.TermTypeDisp = 0
                            regenerateCheckListItems = True
                        End If
                    ElseIf dr.Cells("TYPE TERM@DISP").Value <> pPipe.TermTypeDisp Then
                        If pPipe.TermTypeDisp <= 0 Then
                            dr.Cells("TYPE TERM@DISP").Value = DBNull.Value
                        Else
                            dr.Cells("TYPE TERM@DISP").Value = pPipe.TermTypeDisp
                        End If
                    End If

                    ' Type Term @ Tank
                    If vListTypeTermTank.FindByDataValue(dr.Cells("TYPE TERM@TANK").Value) Is Nothing Then
                        dr.Cells("TYPE TERM@TANK").Value = DBNull.Value
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pPipe.TermTypeTank <> 0 Then
                                pPipe.TermTypeTank = 0
                                regenerateCheckListItems = True
                            End If
                        End If
                    ElseIf dr.Cells("TYPE TERM@TANK").Value Is DBNull.Value Then
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            pPipe.TermTypeTank = 0
                            regenerateCheckListItems = True
                        End If
                    ElseIf dr.Cells("TYPE TERM@TANK").Value <> pPipe.TermTypeTank Then
                        If pPipe.TermTypeTank <= 0 Then
                            dr.Cells("TYPE TERM@TANK").Value = DBNull.Value
                        Else
                            dr.Cells("TYPE TERM@TANK").Value = pPipe.TermTypeTank
                        End If
                    End If

                    If Not vListTypeTermDisp.FindByDataValue(dr.Cells("TYPE TERM@DISP").Value) Is Nothing Then
                        If vListTypeTermDisp.FindByDataValue(dr.Cells("TYPE TERM@DISP").Value).DisplayText.IndexOf("Coated/Wrapped Cathodically Protected") > -1 Then
                            dr.Cells("DISP. TERM CP").Hidden = False
                            ' Dispenser Term CP
                            If vListTDispTermCP.FindByDataValue(dr.Cells("DISP. TERM CP").Value) Is Nothing Then
                                dr.Cells("DISP. TERM CP").Value = DBNull.Value
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    If pPipe.TermCPTypeDisp <> 0 Then
                                        pPipe.TermCPTypeDisp = 0
                                        regenerateCheckListItems = True
                                    End If
                                End If
                            ElseIf dr.Cells("DISP. TERM CP").Value Is DBNull.Value Then
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    pPipe.TermCPTypeDisp = 0
                                    regenerateCheckListItems = True
                                End If
                            ElseIf dr.Cells("DISP. TERM CP").Value <> pPipe.TermCPTypeDisp Then
                                If pPipe.TermCPTypeDisp <= 0 Then
                                    dr.Cells("DISP. TERM CP").Value = DBNull.Value
                                Else
                                    dr.Cells("DISP. TERM CP").Value = pPipe.TermCPTypeDisp
                                End If
                            End If
                        Else
                            dr.Cells("DISP. TERM CP").Hidden = True
                            'TermDispCPTypeCount += 1
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If pPipe.TermCPTypeDisp <> 0 Then
                                    pPipe.TermCPTypeDisp = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        End If
                    Else
                        dr.Cells("DISP. TERM CP").Hidden = True
                        'TermDispCPTypeCount += 1
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pPipe.TermCPTypeDisp <> 0 Then
                                pPipe.TermCPTypeDisp = 0
                                regenerateCheckListItems = True
                            End If
                        End If
                    End If

                    If Not vListTypeTermTank.FindByDataValue(dr.Cells("TYPE TERM@TANK").Value) Is Nothing Then
                        If vListTypeTermTank.FindByDataValue(dr.Cells("TYPE TERM@TANK").Value).DisplayText.IndexOf("Coated/Wrapped Cathodically Protected") > -1 Then
                            dr.Cells("TANK TERM CP").Hidden = False
                            ' Tank Term CP
                            If vListTankTermCP.FindByDataValue(dr.Cells("TANK TERM CP").Value) Is Nothing Then
                                dr.Cells("TANK TERM CP").Value = DBNull.Value
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    If pPipe.TermCPTypeTank <> 0 Then
                                        pPipe.TermCPTypeTank = 0
                                        regenerateCheckListItems = True
                                    End If
                                End If
                            ElseIf dr.Cells("TANK TERM CP").Value Is DBNull.Value Then
                                If bolUpdateObject And Not bolCheckListPrinting Then
                                    pPipe.TermCPTypeTank = 0
                                    regenerateCheckListItems = True
                                End If
                            ElseIf dr.Cells("TANK TERM CP").Value <> pPipe.TermCPTypeTank Then
                                If pPipe.TermCPTypeTank <= 0 Then
                                    dr.Cells("TANK TERM CP").Value = DBNull.Value
                                Else
                                    dr.Cells("TANK TERM CP").Value = pPipe.TermCPTypeTank
                                End If
                            End If
                        Else
                            dr.Cells("TANK TERM CP").Hidden = True
                            'TermTankCPTypeCount += 1
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If pPipe.TermCPTypeTank <> 0 Then
                                    pPipe.TermCPTypeTank = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        End If
                    Else
                        dr.Cells("TANK TERM CP").Hidden = True
                        'TermTankCPTypeCount += 1
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If pPipe.TermCPTypeTank <> 0 Then
                                pPipe.TermCPTypeTank = 0
                                regenerateCheckListItems = True
                            End If
                        End If
                    End If

                    ' Enable Field              Condition
                    ' Pipe Term CP Installed    Pipe Term Type at Tank / Disp = 'Coated/Wrapped Cathodically Protected'
                    ' Pipe Term CP Last Tested  Pipe Term Type at Tank / Disp = 'Coated/Wrapped Cathodically Protected'
                    Dim bolHideTermCPTested As Boolean = True

                    If Not vListTypeTermDisp.FindByDataValue(dr.Cells("TYPE TERM@DISP").Value) Is Nothing Then
                        If vListTypeTermDisp.FindByDataValue(dr.Cells("TYPE TERM@DISP").Value).DisplayText.IndexOf("Coated/Wrapped Cathodically Protected") > -1 Then
                            'dr.Cells("TERM CP TESTED").Hidden = False
                            bolHideTermCPTested = False
                            'If Not oInspection.CAPDatesEntered Then
                            '    dr.Cells("CP TESTED").Value = DBNull.Value
                            '    pPipe.TermCPLastTested = dtnulldate
                            'End If
                            'Else
                            'dr.Cells("TERM CP TESTED").Hidden = True
                            'TermCPTestedCount += 1
                        End If
                    End If

                    If Not vListTypeTermTank.FindByDataValue(dr.Cells("TYPE TERM@TANK").Value) Is Nothing Then
                        If vListTypeTermTank.FindByDataValue(dr.Cells("TYPE TERM@TANK").Value).DisplayText.IndexOf("Coated/Wrapped Cathodically Protected") > -1 Then
                            '    dr.Cells("TERM CP TESTED").Hidden = False
                            bolHideTermCPTested = False
                            '    If Not oInspection.CAPDatesEntered Then
                            '        dr.Cells("CP TESTED").Value = DBNull.Value
                            '        pPipe.TermCPLastTested = dtnulldate
                            '    End If
                            'Else
                            '    dr.Cells("TERM CP TESTED").Hidden = True
                            'TermCPTestedCount += 1
                        End If
                    End If

                    If bolHideTermCPTested Then
                        dr.Cells("TERM CP INSTALLED").Hidden = True
                        dr.Cells("TERM CP TESTED").Hidden = True
                        If bolUpdateObject And Not bolCheckListPrinting Then
                            If Date.Compare(pPipe.TermCPLastTested, dtNullDate) <> 0 Then
                                pPipe.TermCPLastTested = dtNullDate
                                regenerateCheckListItems = True
                            End If
                            If Date.Compare(pPipe.TermCPInstalledDate, dtNullDate) <> 0 Then
                                pPipe.TermCPInstalledDate = dtNullDate
                                regenerateCheckListItems = True
                            End If
                        End If
                    Else
                        dr.Cells("TERM CP INSTALLED").Hidden = False
                        dr.Cells("TERM CP TESTED").Hidden = False
                        If Not oInspection.CAPDatesEntered Then
                            dr.Cells("TERM CP TESTED").Value = DBNull.Value
                            'Manju
                            If bolUpdateObject And Not bolCheckListPrinting Then
                                If Date.Compare(pPipe.TermCPLastTested, dtNullDate) <> 0 Then
                                    pPipe.TermCPLastTested = dtNullDate
                                    regenerateCheckListItems = True
                                End If
                                'Else
                                '    pPipe.TermCPLastTested = dtNullDate
                            End If
                        End If
                    End If

                    ' make the row readonly if checklist is opened in readonly mode
                    If [readOnly] Then
                        dr.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    End If
                End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetupTankPipeLeakRow(ByRef ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            If Not bolReadOnly Then
                If Not ugRow.Cells("Well#").Value Is DBNull.Value Then
                    If ugRow.Cells("Well#").Value = 0 Then
                        ugRow.Cells("Well#").Value = DBNull.Value
                    End If
                End If
                'If Not ugRow.Cells("Well Depth").Value Is DBNull.Value Then
                '    If ugRow.Cells("Well Depth").Value = 0 Then
                '        ugRow.Cells("Well Depth").Value = DBNull.Value
                '    End If
                'End If
                'If Not ugRow.Cells("Depth to" + vbCrLf + "Water").Value Is DBNull.Value Then
                '    If ugRow.Cells("Depth to" + vbCrLf + "Water").Value = 0 Then
                '        ugRow.Cells("Depth to" + vbCrLf + "Water").Value = DBNull.Value
                '    End If
                'End If
                'If Not ugRow.Cells("Depth to" + vbCrLf + "Slots").Value Is DBNull.Value Then
                '    If ugRow.Cells("Depth to" + vbCrLf + "Slots").Value = 0 Then
                '        ugRow.Cells("Depth to" + vbCrLf + "Slots").Value = DBNull.Value
                '    End If
                'End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetupRectifierRow(ByRef ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            If Not ugRow.Cells("Volts").Value Is DBNull.Value Then
                If ugRow.Cells("Volts").Value = 0.0 Then
                    ugRow.Cells("Volts").Value = DBNull.Value
                End If
            End If
            If Not ugRow.Cells("Amps").Value Is DBNull.Value Then
                If ugRow.Cells("Amps").Value = 0.0 Then
                    ugRow.Cells("Amps").Value = DBNull.Value
                End If
            End If
            If Not ugRow.Cells("Hours").Value Is DBNull.Value Then
                If ugRow.Cells("Hours").Value = 0.0 Then
                    ugRow.Cells("Hours").Value = DBNull.Value
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Private Sub ugTankValidation(ByRef errStrLocal As String, ByVal e As Infragistics.Win.UltraWinGrid.UltraGridRow)
    '    Try
    '        ' tank req fields validation
    '        Dim substanceUsedOilCount As Int64 = 0
    '        ' 12
    '        If Date.Compare(oTank.DateInstalledTank, dtNullDate) = 0 Then
    '            errStrLocal += "INSTALLED" + ", "
    '        End If
    '        ' 15
    '        If oTank.TankMatDesc = 0 Then
    '            errStrLocal += "MATERIALS" + ", "
    '        End If
    '        ' 16
    '        If oTank.TankModDesc = 0 Then
    '            errStrLocal += "OPTIONS" + ", "
    '        End If
    '        ' 17
    '        If e.Cells("CP TYPE").Hidden = False Then
    '            If oTank.TankCPType = 0 Then
    '                errStrLocal += "CP TYPE" + ", "
    '            End If
    '        End If
    '        ' 18
    '        If oTank.TankLD = 0 Then
    '            errStrLocal += "LEAK DETECTION" + ", "
    '        End If
    '        ' 20
    '        If oTank.OverFillType = 0 Then
    '            errStrLocal += "OVERFILL TYPE" + ", "
    '        End If
    '        ' 21
    '        If e.Cells("LINED").Hidden = False Then
    '            If Date.Compare(oTank.LinedInteriorInstallDate, dtNullDate) = 0 Then
    '                errStrLocal += "LINED" + ", "
    '            End If
    '        End If
    '        ' 22
    '        If e.Cells("LINING INSPECTED").Hidden = False Then
    '            For Each clInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.InspectionInfo.ChecklistMasterCollection.Values
    '                If clInfo.CheckListItemNumber = "3.5.3" Then
    '                    If clInfo.Show Then
    '                        For Each resp As MUSTER.Info.InspectionResponsesInfo In oInspection.InspectionInfo.ResponsesCollection.Values
    '                            If resp.QuestionID = clInfo.ID Then
    '                                If resp.Response = 1 Then
    '                                    If Date.Compare(oTank.LinedInteriorInspectDate, dtNullDate) = 0 Then
    '                                        errStrLocal += "LINING INSPECTED" + ", "
    '                                        Exit For
    '                                    End If
    '                                Else
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next
    '                        Exit For
    '                    Else
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        End If
    '        ' 23
    '        If e.Cells("PTT").Hidden = False Then
    '            For Each clInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.InspectionInfo.ChecklistMasterCollection.Values
    '                If clInfo.CheckListItemNumber = "4.3.2" Then
    '                    If clInfo.Show Then
    '                        For Each resp As MUSTER.Info.InspectionResponsesInfo In oInspection.InspectionInfo.ResponsesCollection.Values
    '                            If resp.QuestionID = clInfo.ID Then
    '                                If resp.Response = 1 Then
    '                                    If Date.Compare(oTank.TTTDate, dtNullDate) = 0 Then
    '                                        errStrLocal += "PTT" + ", "
    '                                        Exit For
    '                                    End If
    '                                Else
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next
    '                        Exit For
    '                    Else
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        End If
    '        ' 24
    '        'If e.Cells("CP INSTALLED").Hidden = False Then
    '        ' TODO
    '        'End If
    '        ' 25
    '        If e.Cells("CP TESTED").Hidden = False Then
    '            For Each clInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.InspectionInfo.ChecklistMasterCollection.Values
    '                If clInfo.CheckListItemNumber = "3.5.2" Then
    '                    If clInfo.Show Then
    '                        For Each resp As MUSTER.Info.InspectionResponsesInfo In oInspection.InspectionInfo.ResponsesCollection.Values
    '                            If resp.QuestionID = clInfo.ID Then
    '                                If resp.Response = 1 Then
    '                                    If Date.Compare(oTank.LastTCPDate, dtNullDate) = 0 Then
    '                                        errStrLocal += "CP TESTED" + ", "
    '                                        Exit For
    '                                    End If
    '                                Else
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next
    '                        Exit For
    '                    Else
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugPipeValidation(ByRef errStrLocal As String, ByVal e As Infragistics.Win.UltraWinGrid.UltraGridRow)
    '    Try
    '        ' pipe req fields validation
    '        ' 27
    '        If oPipe.PipeStatusDesc = 0 Then
    '            errStrLocal += "STATUS" + ", "
    '        End If
    '        ' 28
    '        If Date.Compare(oPipe.PipeInstallDate, dtNullDate) = 0 Then
    '            errStrLocal += "INSTALLED" + ", "
    '        End If
    '        ' 29
    '        If oPipe.PipeTypeDesc = 0 Then
    '            errStrLocal += "TYPE OF SYSTEM" + ", "
    '        End If
    '        ' 30
    '        If oPipe.PipeMatDesc = 0 Then
    '            errStrLocal += "MATERIALS" + ", "
    '        End If
    '        ' 31
    '        If oPipe.PipeModDesc = 0 Then
    '            errStrLocal += "OPTIONS" + ", "
    '        End If
    '        ' 33
    '        If e.Cells("CP TYPE").Hidden = False Then
    '            If oPipe.PipeCPType = 0 Then
    '                errStrLocal += "CP TYPE" + ", "
    '            End If
    '        End If
    '        ' 34
    '        If e.Cells("PRI. LEAK DETECTION").Hidden = False Then
    '            If oPipe.PipeLD = 0 Then
    '                errStrLocal += "PRI. LEAK DETECTION" + ", "
    '            End If
    '        End If
    '        ' 35
    '        If e.Cells("SEC. LEAK DETECTION").Hidden = False Then
    '            If oPipe.ALLDType = 0 Then
    '                errStrLocal += "SEC. LEAK DETECTION" + ", "
    '            End If
    '        End If
    '        ' 36
    '        If e.Cells("ALLD TESTED").Hidden = False Then
    '            For Each clInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.InspectionInfo.ChecklistMasterCollection.Values
    '                If clInfo.CheckListItemNumber = "5.9.4" Then
    '                    If clInfo.Show Then
    '                        For Each resp As MUSTER.Info.InspectionResponsesInfo In oInspection.InspectionInfo.ResponsesCollection.Values
    '                            If resp.QuestionID = clInfo.ID Then
    '                                If resp.Response = 1 Then
    '                                    If Date.Compare(oPipe.ALLDTestDate, dtNullDate) = 0 Then
    '                                        errStrLocal += "ALLD TESTED" + ", "
    '                                        Exit For
    '                                    End If
    '                                Else
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next
    '                        Exit For
    '                    Else
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        End If
    '        ' 37
    '        If e.Cells("PTT").Hidden = False Then
    '            For Each clInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.InspectionInfo.ChecklistMasterCollection.Values
    '                If clInfo.CheckListItemNumber = "5.3.1" Or clInfo.CheckListItemNumber = "5.3.2" Then
    '                    If clInfo.Show Then
    '                        For Each resp As MUSTER.Info.InspectionResponsesInfo In oInspection.InspectionInfo.ResponsesCollection.Values
    '                            If resp.QuestionID = clInfo.ID Then
    '                                If resp.Response = 1 Then
    '                                    If Date.Compare(oPipe.LTTDate, dtNullDate) = 0 Then
    '                                        errStrLocal += "PTT" + ", "
    '                                        Exit For
    '                                    End If
    '                                Else
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next
    '                        Exit For
    '                    Else
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        End If
    '        ' 39
    '        If e.Cells("CP TESTED").Hidden = False Then
    '            For Each clInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.InspectionInfo.ChecklistMasterCollection.Values
    '                If clInfo.CheckListItemNumber = "3.6.2" Then
    '                    If clInfo.Show Then
    '                        For Each resp As MUSTER.Info.InspectionResponsesInfo In oInspection.InspectionInfo.ResponsesCollection.Values
    '                            If resp.QuestionID = clInfo.ID Then
    '                                If resp.Response = 1 Then
    '                                    If Date.Compare(oPipe.PipeCPTest, dtNullDate) = 0 Then
    '                                        errStrLocal += "CP TESTED" + ", "
    '                                        Exit For
    '                                    End If
    '                                Else
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next
    '                        Exit For
    '                    Else
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugTermValidation(ByRef errStrLocal As String, ByVal e As Infragistics.Win.UltraWinGrid.UltraGridRow)
    '    Try
    '        ' term req fields validation
    '        ' 42
    '        If oPipe.TermTypeTank = 0 Then
    '            errStrLocal += "TYPE TERM@TANK" + ", "
    '        End If
    '        ' 43
    '        If oPipe.TermTypeDisp = 0 Then
    '            errStrLocal += "TYPE TERM@DISP" + ", "
    '        End If
    '        ' 44
    '        If oPipe.TermCPTypeTank = 0 Then
    '            errStrLocal += "TANK TERM CP" + ", "
    '        End If
    '        ' 45
    '        If oPipe.TermCPTypeDisp = 0 Then
    '            errStrLocal += "DISP. TERM CP" + ", "
    '        End If
    '        ' 46
    '        If e.Cells("TERM CP TESTED").Hidden = False Then
    '            For Each clInfo As MUSTER.Info.InspectionChecklistMasterInfo In oInspection.InspectionInfo.ChecklistMasterCollection.Values
    '                If clInfo.CheckListItemNumber = "3.7.5" Then
    '                    If clInfo.Show Then
    '                        For Each resp As MUSTER.Info.InspectionResponsesInfo In oInspection.InspectionInfo.ResponsesCollection.Values
    '                            If resp.QuestionID = clInfo.ID Then
    '                                If resp.Response = 1 Then
    '                                    If Date.Compare(oPipe.TermCPLastTested, dtNullDate) = 0 Then
    '                                        errStrLocal += "TERM CP TESTED" + ", "
    '                                        Exit For
    '                                    End If
    '                                Else
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next
    '                        Exit For
    '                    Else
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub ValidateTankPipeTerm(ByRef strErr As String, Optional ByRef allOptionalValidations As Boolean = True)
        Dim row As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim strTankPipeTermErr, strTankErrRow, strPipeErrRow, strTermErrRow As String
        Dim id As String
        Try
            'strTankPipeTermErr = "The following field(s) are required"
            oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
            oPipe = oTank.Pipes
            For Each row In ugTanks.Rows
                If row.Appearance.BackColor.Name <> Color.LightGray.Name Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(row.Cells("TANK_ID").Value.ToString)
                    strTankErrRow = String.Empty
                    ' 2197 validate cap dates only if facility is cap candidate
                    oTank.ValidateData(oInspection.CheckListMaster.Owner.Facility.CAPCandidate, "Inspection", strTankErrRow, True)
                    If Not strTankErrRow.StartsWith("Optional") And strTankErrRow <> String.Empty Then allOptionalValidations = False
                    'ugTankValidation(strTankErrRow, row)
                    If strTankErrRow.Length > 0 Then
                        'strTankPipeTermErr += vbCrLf + "Tank #:" + row.Cells("Tank #").Text + " - " + strTankErrRow.Trim.TrimEnd(",")
                        strErr += vbCrLf + vbCrLf + "Tank #:" + row.Cells("Tank #").Text + " - " + strTankErrRow.Trim.TrimEnd(",")
                    End If
                End If
            Next
            For Each row In ugPipes.Rows
                id = row.Cells("TANK_ID").Value.ToString + "|" + _
                        row.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + _
                        row.Cells("PIPE_ID").Value.ToString
                oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(row.Cells("TANK_ID").Value.ToString)
                oPipe.Pipe = oTank.TankInfo.pipesCollection.Item(id)
                strPipeErrRow = String.Empty
                ' 2197 validate cap dates only if facility is cap candidate
                oPipe.ValidateData(oInspection.CheckListMaster.Owner.Facility.CAPCandidate, "Inspection", strPipeErrRow, True)
                If Not strPipeErrRow.StartsWith("Optional") And strPipeErrRow <> String.Empty Then allOptionalValidations = False
                'ugPipeValidation(strPipeErrRow, row)
                If strPipeErrRow.Length > 0 Then
                    'strTankPipeTermErr += vbCrLf + "Pipe #:" + row.Cells("Pipe #").Text + " - " + strPipeErrRow.Trim.TrimEnd(",")
                    strErr += vbCrLf + vbCrLf + "Pipe #:" + row.Cells("Pipe #").Text + " - " + strPipeErrRow.Trim.TrimEnd(",")
                End If
            Next
            'For Each row In ugTerminations.Rows
            '    id = row.Cells("TANK_ID").Value.ToString + "|" + _
            '            row.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + _
            '            row.Cells("PIPE_ID").Value.ToString
            '    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(row.Cells("TANK_ID").Value.ToString)
            '    oPipe.Pipe = oTank.TankInfo.pipesCollection.Item(id)
            '    strTermErrRow = String.Empty
            '    ugTermValidation(strTermErrRow, row)
            '    If strTankErrRow.Length > 0 Then
            '        'strTankPipeTermErr += vbCrLf + "Pipe #:" + row.Cells("Pipe #").Text + " - " + strTermErrRow.Trim.TrimEnd(",")
            '        strErr += vbCrLf + "Pipe #:" + row.Cells("Pipe #").Text + " - " + strTermErrRow.Trim.TrimEnd(",")
            '    End If
            'Next
            If strErr.Length > 0 Then
                strErr = IIf(allOptionalValidations, "The following field(s) are optional", "The following field(s) are required / optional") + strErr
            End If
            'strErr = IIf(strTankPipeTermErr <> "The following field(s) are required for", strTankPipeTermErr, String.Empty)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Private Sub ValidateResponses1(ByRef strErr As String)
    '    Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    '    Dim str As String = String.Empty
    '    Dim dtProcessed As Boolean = False
    '    Try
    '        btnReg.PerformClick()
    '        For Each ugRow In ugReg.Rows
    '            If ugRow.Cells("HEADER").Value = False Then
    '                If ugRow.Cells.Exists("Line#") Then
    '                    If ugRow.Cells("RESPONSE").Value = -1 Then
    '                        str += ugRow.Cells("Line#").Value.ToString + ", "
    '                    End If
    '                End If
    '            End If
    '        Next
    '        If str.Length > 0 Then
    '            strErr += str.Trim.TrimEnd(",") + vbCrLf
    '            str = String.Empty
    '        End If

    '        btnSpill.PerformClick()
    '        For Each ugRow In ugSpill.Rows
    '            If ugRow.Cells("HEADER").Value = False Then
    '                If ugRow.Cells.Exists("Line#") Then
    '                    If ugRow.Cells("RESPONSE").Value = -1 Then
    '                        str += ugRow.Cells("Line#").Value.ToString + ", "
    '                    End If
    '                End If
    '            End If
    '        Next
    '        If str.Length > 0 Then
    '            strErr += str.Trim.TrimEnd(",") + vbCrLf
    '            str = String.Empty
    '        End If

    '        btnCp.PerformClick()
    '        For Each ugRow In ugCP.Rows
    '            If ugRow.Band.Index = 0 Then
    '                If ugRow.Cells("HEADER").Value = False Then
    '                    If ugRow.Cells.Exists("Line#") Then
    '                        If ugRow.Cells("RESPONSE").Value = -1 Then
    '                            str += ugRow.Cells("Line#").Value.ToString + ", "
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next
    '        If str.Length > 0 Then
    '            strErr += str.Trim.TrimEnd(",") + vbCrLf
    '            str = String.Empty
    '        End If

    '        btnCp.PerformClick()
    '        For Each ugRow In ugCP.Rows
    '            If ugRow.Cells("HEADER").Value = False Then
    '                If ugRow.Cells.Exists("Line#") Then
    '                    If ugRow.Cells("RESPONSE").Value = -1 Then
    '                        str += ugRow.Cells("Line#").Value.ToString + ", "
    '                    End If
    '                End If
    '            End If
    '        Next
    '        If str.Length > 0 Then
    '            strErr += str.Trim.TrimEnd(",") + vbCrLf
    '            str = String.Empty
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub ValidateResponses(ByRef strErr As String)
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim dr As DataRow
        Dim dt, dTable As DataTable
        Dim ds As DataSet
        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
        Dim resp As MUSTER.Info.InspectionResponsesInfo
        Dim cp As MUSTER.Info.InspectionCPReadingsInfo
        Dim well As MUSTER.Info.InspectionMonitorWellsInfo
        Dim validateDone As Boolean = False
        Dim ugName As String = String.Empty
        Dim str As String = String.Empty
        Dim dtProcessed As Boolean = False
        Dim rowIndex As Integer
        Dim arrUGName(8) As String
        Dim bolSkipCP As Boolean = False
        Try
            arrUGName(0) = ugReg.Name
            arrUGName(1) = ugSpill.Name
            arrUGName(2) = ugCP.Name
            arrUGName(3) = ugTankLeak.Name
            arrUGName(4) = ugPipeLeak.Name
            arrUGName(5) = ugCatLeak.Name
            arrUGName(6) = ugVisual.Name
            arrUGName(7) = ugTOS.Name


            For i As Integer = 0 To arrUGName.GetUpperBound(0) - 1
                ugName = arrUGName(i)
                Select Case ugName
                    Case ugReg.Name
                        ds = New DataSet
                        dTable = oInspection.CheckListMaster.RegTable
                        ds.Tables.Add(dTable)
                    Case ugSpill.Name
                        ds = New DataSet
                        dTable = oInspection.CheckListMaster.SpillTable
                        ds.Tables.Add(dTable)
                    Case ugCP.Name
                        ds = New DataSet
                        ds = oInspection.CheckListMaster.CPTable
                        ' #2903 validation should not include hidden rows
                        bol354Hidden = True
                        bol363Hidden = True
                        bol376Hidden = True
                        ' tank
                        If Not ds.Tables("CPTankInspectorTested") Is Nothing Then
                            If ds.Tables("CPTankInspectorTested").Rows.Count > 0 Then
                                If ds.Tables("CPTankInspectorTested").Rows(0)("Yes") Then
                                    bol354Hidden = False
                                End If
                            End If
                        End If
                        ' pipe
                        If Not ds.Tables("CPPipeInspectorTested") Is Nothing Then
                            If ds.Tables("CPPipeInspectorTested").Rows.Count > 0 Then
                                If ds.Tables("CPPipeInspectorTested").Rows(0)("Yes") Then
                                    bol363Hidden = False
                                End If
                            End If
                        End If
                        ' term
                        If Not ds.Tables("CPTermInspectorTested") Is Nothing Then
                            If ds.Tables("CPTermInspectorTested").Rows.Count > 0 Then
                                If ds.Tables("CPTermInspectorTested").Rows(0)("Yes") Then
                                    bol376Hidden = False
                                End If
                            End If
                        End If
                    Case ugTankLeak.Name
                        ds = New DataSet
                        ds = oInspection.CheckListMaster.TankLeakTable
                    Case ugPipeLeak.Name
                        ds = New DataSet
                        ds = oInspection.CheckListMaster.PipeLeakTable
                    Case ugCatLeak.Name
                        ds = New DataSet
                        dTable = oInspection.CheckListMaster.CATLeakTable
                        ds.Tables.Add(dTable)
                    Case ugVisual.Name
                        ds = New DataSet
                        dTable = oInspection.CheckListMaster.VisualTable
                        ds.Tables.Add(dTable)
                    Case ugTOS.Name
                        ds = New DataSet
                        dTable = oInspection.CheckListMaster.TOSTable
                        ds.Tables.Add(dTable)
                End Select
                For Each dt In ds.Tables
                    dtProcessed = False
                    If dt.Columns.Contains("Pass") Then
                        dtProcessed = True
                        For rowIndex = 0 To dt.Rows.Count - 1
                            bolSkipCP = False
                            cp = oInspection.InspectionInfo.CPReadingsCollection.Item(dt.DefaultView.Item(rowIndex)("ID"))
                            ' to determine tank/pipe/term cp reading
                            If cp.QuestionID = oInspection.CheckListMaster.QuestionIDofCPItem354 And oInspection.CheckListMaster.QuestionIDofCPItem354 > 0 Then
                                bolSkipCP = bol354Hidden
                            ElseIf cp.QuestionID = oInspection.CheckListMaster.QuestionIDofCPItem363 And oInspection.CheckListMaster.QuestionIDofCPItem363 > 0 Then
                                bolSkipCP = bol363Hidden
                            ElseIf cp.QuestionID = oInspection.CheckListMaster.QuestionIDofCPItem376 And oInspection.CheckListMaster.QuestionIDofCPItem376 > 0 Then
                                bolSkipCP = bol376Hidden
                            End If
                            If Not bolSkipCP Then
                                If cp.PassFailIncon = -1 Then
                                    str += dt.DefaultView.Item(rowIndex)("Line#").ToString + ", "
                                End If
                            End If
                        Next
                    End If
                    If Not dtProcessed Then
                        If dt.Columns.Contains("Surface Sealed" + vbCrLf + "No") Then
                            dtProcessed = True
                            For rowIndex = 0 To dt.Rows.Count - 1
                                well = oInspection.InspectionInfo.MonitorWellsCollection.Item(dt.DefaultView.Item(rowIndex)("ID"))
                                If well.SurfaceSealed = -1 Or well.WellCaps = -1 Then
                                    str += dt.DefaultView.Item(rowIndex)("Line#").ToString + ", "
                                End If
                            Next
                        End If
                    End If
                    If Not dtProcessed Then
                        If dt.Columns.Contains("Line#") Then
                            For rowIndex = 0 To dt.Rows.Count - 1
                                checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(dt.DefaultView.Item(rowIndex)("QUESTION_ID"))
                                If Not checkList.Header Then
                                    resp = oInspection.InspectionInfo.ResponsesCollection.Item(dt.DefaultView.Item(rowIndex)("ID"))
                                    If resp.Response = -1 Then
                                        str += dt.DefaultView.Item(rowIndex)("Line#").ToString + ", "
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If str.Length > 0 Then
                        strErr += str.Trim.TrimEnd(",") + vbCrLf
                        str = String.Empty
                    End If
                Next
            Next
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Function ValidateAddress() As Boolean
        Try
            AddressForm = New Address(UIUtilsGen.EntityTypes.Facility, oInspection.CheckListMaster.Owner.Facilities.FacilityAddresses, "Facility", oInspection.CheckListMaster.Owner.AddressId)
            AddressForm.ShowFIPS = False
            Return AddressForm.ValidateData
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Sub PrintCheckList(Optional ByVal strHeaderText As String = "Printing progress")
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugcell As Infragistics.Win.UltraWinGrid.UltraGridCell
        Dim dsTanks, dsPipes, dsTerms As DataSet
        Dim dtTank, dtPipe, dtTerm As DataTable
        Dim dtCol As DataColumn
        Dim strColName As String
        Dim dr As DataRow
        Try
            bolCheckListPrinting = True
            ' show progress bar
            frmCLProgress = New CheckListProgress
            frmCLProgress.HeaderText = strHeaderText
            frmCLProgress.Show()
            UIUtilsGen.Delay(, 0.5)

            dsTanks = New DataSet
            dsPipes = New DataSet
            dsTerms = New DataSet
            Dim bolLoadingLocal As Boolean = bolLoading
            Try
                bolLoading = True
                PrepTankPipeTerm()
            Finally
                bolLoading = bolLoadingLocal
            End Try
            For Each ugrow In ugTanks.Rows
                ' create new table if table already not present
                If dsTanks.Tables(ugrow.Cells("TANK_ID").Text + "|" + ugrow.Cells("COMPARTMENT #").Text) Is Nothing Then
                    dtTank = New DataTable
                    For Each ugcell In ugrow.Cells
                        If ugcell.Hidden = False And ugcell.Column.Hidden = False Then
                            Select Case ugcell.Column.Key.Trim.ToUpper
                                Case "TANK #"
                                    dtTank.Columns.Add("Tank #")
                                Case "COMPARTMENT #"
                                    dtTank.Columns.Add("Comp #")
                                Case "STATUS"
                                    dtTank.Columns.Add("Status")
                                Case "INSTALLED"
                                    dtTank.Columns.Add("Installed")
                                Case "PLACED IN SERVICE"
                                    dtTank.Columns.Add("Placed In Service")
                                Case "SIZE"
                                    dtTank.Columns.Add("Size")
                                Case "CONTENTS"
                                    dtTank.Columns.Add("Contents")
                                Case "FUEL TYPE"
                                    dtTank.Columns.Add("Fuel Type")
                                Case "MATERIALS"
                                    dtTank.Columns.Add("Material of Construction")
                                Case "OPTIONS"
                                    dtTank.Columns.Add("Construction Options")
                                Case "CP TYPE"
                                    dtTank.Columns.Add("CP Type")
                                Case "LEAK DETECTION"
                                    dtTank.Columns.Add("Leak Detection")
                                Case "OVERFILL TYPE"
                                    dtTank.Columns.Add("Overfill Type")
                                Case "LINED"
                                    dtTank.Columns.Add("Lined")
                                Case "LINING INSPECTED"
                                    dtTank.Columns.Add("Lining Inspected")
                                Case "PTT"
                                    dtTank.Columns.Add("PTT")
                                Case "CP INSTALLED"
                                    dtTank.Columns.Add("CP Installed")
                                Case "CP TESTED"
                                    dtTank.Columns.Add("CP Tested")
                            End Select
                        End If
                    Next
                    dr = dtTank.NewRow
                    For Each dtCol In dtTank.Columns
                        Select Case dtCol.ColumnName
                            Case "Tank #"
                                strColName = "TANK #"
                            Case "Comp #"
                                strColName = "COMPARTMENT #"
                            Case "Status"
                                strColName = "STATUS"
                            Case "Installed"
                                strColName = "INSTALLED"
                            Case "PLACED IN SERVICE"
                                strColName = "Placed In Service"
                            Case "Size"
                                strColName = "SIZE"
                            Case "Contents"
                                strColName = "CONTENTS"
                            Case "Fuel Type"
                                strColName = "FUEL TYPE"
                            Case "Material of Construction"
                                strColName = "MATERIALS"
                            Case "Construction Options"
                                strColName = "OPTIONS"
                            Case "CP Type"
                                strColName = "CP TYPE"
                            Case "Leak Detection"
                                strColName = "LEAK DETECTION"
                            Case "Overfill Type"
                                strColName = "OVERFILL TYPE"
                            Case "Lined"
                                strColName = "LINED"
                            Case "Lining Inspected"
                                strColName = "LINING INSPECTED"
                            Case "PTT"
                                strColName = "PTT"
                            Case "CP Installed"
                                strColName = "CP INSTALLED"
                            Case "CP Tested"
                                strColName = "CP TESTED"
                        End Select
                        dr(dtCol.ColumnName) = ugrow.Cells(strColName).Text
                    Next
                    dtTank.Rows.Add(dr)
                    dtTank.TableName = ugrow.Cells("TANK_ID").Text + "|" + ugrow.Cells("COMPARTMENT #").Text
                    dsTanks.Tables.Add(dtTank)
                End If
            Next
            frmCLProgress.ProgressBarValue += 10
            For Each ugrow In ugPipes.Rows
                If dsPipes.Tables(ugrow.Cells("PIPE_ID").Text) Is Nothing Then
                    dtPipe = New DataTable
                    For Each ugcell In ugrow.Cells
                        If ugcell.Hidden = False And ugcell.Column.Hidden = False Then
                            Select Case ugcell.Column.Key.Trim.ToUpper
                                Case "PIPE #"
                                    dtPipe.Columns.Add("Pipe #")
                                Case "STATUS"
                                    dtPipe.Columns.Add("Status")
                                Case "INSTALLED"
                                    dtPipe.Columns.Add("Installed")
                                Case "TYPE OF SYSTEM"
                                    dtPipe.Columns.Add("Pipe Type")
                                Case "MATERIALS"
                                    dtPipe.Columns.Add("Material of Construction")
                                Case "OPTIONS"
                                    dtPipe.Columns.Add("Construction Options")
                                Case "BRAND"
                                    dtPipe.Columns.Add("Brand")
                                Case "CP TYPE"
                                    dtPipe.Columns.Add("CP Type")
                                Case "PRI. LEAK DETECTION"
                                    dtPipe.Columns.Add("Primary LD")
                                Case "SEC. LEAK DETECTION"
                                    dtPipe.Columns.Add("Secondary LD")
                                Case "ALLD TESTED"
                                    dtPipe.Columns.Add("ALLD Tested")
                                Case "PTT"
                                    dtPipe.Columns.Add("PTT")
                                Case "CP INSTALLED"
                                    dtPipe.Columns.Add("CP Installed")
                                Case "CP TESTED"
                                    dtPipe.Columns.Add("CP Tested")
                            End Select
                        End If
                    Next
                    dr = dtPipe.NewRow
                    For Each dtCol In dtPipe.Columns
                        Select Case dtCol.ColumnName
                            Case "Pipe #"
                                strColName = "PIPE #"
                            Case "Status"
                                strColName = "STATUS"
                            Case "Installed"
                                strColName = "INSTALLED"
                            Case "Pipe Type"
                                strColName = "TYPE OF SYSTEM"
                            Case "Material of Construction"
                                strColName = "MATERIALS"
                            Case "Construction Options"
                                strColName = "OPTIONS"
                            Case "Brand"
                                strColName = "BRAND"
                            Case "CP Type"
                                strColName = "CP TYPE"
                            Case "Primary LD"
                                strColName = "PRI. LEAK DETECTION"
                            Case "Secondary LD"
                                strColName = "SEC. LEAK DETECTION"
                            Case "ALLD Tested"
                                strColName = "ALLD TESTED"
                            Case "PTT"
                                strColName = "PTT"
                            Case "CP Installed"
                                strColName = "CP INSTALLED"
                            Case "CP Tested"
                                strColName = "CP TESTED"
                        End Select
                        dr(dtCol.ColumnName) = ugrow.Cells(strColName).Text
                    Next
                    dtPipe.Rows.Add(dr)
                    dtPipe.TableName = ugrow.Cells("PIPE_ID").Text
                    dsPipes.Tables.Add(dtPipe)
                End If
            Next
            frmCLProgress.ProgressBarValue += 10
            For Each ugrow In ugTerminations.Rows
                If dsTerms.Tables(ugrow.Cells("PIPE_ID").Text) Is Nothing Then
                    dtTerm = New DataTable
                    For Each ugcell In ugrow.Cells
                        If ugcell.Hidden = False And ugcell.Column.Hidden = False Then
                            Select Case ugcell.Column.Key.Trim.ToUpper
                                Case "PIPE #"
                                    dtTerm.Columns.Add("Pipe #")
                                Case "SUMP@TANK"
                                    dtTerm.Columns.Add("Sump @ Tank")
                                Case "SUMP@DISP"
                                    dtTerm.Columns.Add("Sump @ Dispenser")
                                Case "TYPE TERM@TANK"
                                    dtTerm.Columns.Add("Type Term. @ Tank")
                                Case "TYPE TERM@DISP"
                                    dtTerm.Columns.Add("Type Term. @ Dispenser")
                                Case "TANK TERM CP"
                                    dtTerm.Columns.Add("Tank Term. CP")
                                Case "DISP. TERM CP"
                                    dtTerm.Columns.Add("Dispenser Term. CP")
                                Case "TERM CP TESTED"
                                    dtTerm.Columns.Add("Term. CP Tested")
                            End Select
                        End If
                    Next
                    dr = dtTerm.NewRow
                    For Each dtCol In dtTerm.Columns
                        Select Case dtCol.ColumnName
                            Case "Pipe #"
                                strColName = "PIPE #"
                            Case "Sump @ Tank"
                                strColName = "SUMP@TANK"
                            Case "Sump @ Dispenser"
                                strColName = "SUMP@DISP"
                            Case "Type Term. @ Tank"
                                strColName = "TYPE TERM@TANK"
                            Case "Type Term. @ Dispenser"
                                strColName = "TYPE TERM@DISP"
                            Case "Tank Term. CP"
                                strColName = "TANK TERM CP"
                            Case "Dispenser Term. CP"
                                strColName = "DISP. TERM CP"
                            Case "Term. CP Tested"
                                strColName = "TERM CP TESTED"
                        End Select
                        dr(dtCol.ColumnName) = ugrow.Cells(strColName).Text
                    Next
                    dtTerm.Rows.Add(dr)
                    dtTerm.TableName = ugrow.Cells("PIPE_ID").Text
                    dsTerms.Tables.Add(dtTerm)
                End If
            Next
            frmCLProgress.ProgressBarValue += 10
            ltrGen = New MUSTER.BusinessLogic.pLetterGen
            Dim oLetter As New Reg_Letters
            oLetter.GenerateInspectionCheckList(oInspection.CheckListMaster.Owner.Facilities.ID, "Inspection CheckList", "CheckList", "Facility Inspection CheckList", "InspectionCheckList.doc", oInspection, dsTanks, dsPipes, dsTerms, ltrGen, frmCLProgress.ProgressBarValue, GetFlagsForPrintedChecklist())
            ' Delay after closing a word document to resolve RPC Server is Unavailable Issue
            'UIUtilsGen.Delay(, 1)
        Catch ex As Exception
            Throw ex
        Finally
            bolCheckListPrinting = False
            frmCLProgress.Close()
        End Try
    End Sub
    Private Function ValidateCheckList(ByVal bolValidateTankPipeTerm As Boolean, Optional ByVal bolAllowYesNo As Boolean = True) As Boolean
        Dim strTankPipeTermErr As String = String.Empty
        Dim strResponseErr As String = String.Empty
        Dim bolReturnValue As Boolean = False
        Dim bolAllOptionalTankPipeErr As Boolean = True
        Try
            ' validate address
            If Not ValidateAddress() Then
                Exit Function
            End If
            If bolValidateTankPipeTerm Then
                Cursor.Current = Cursors.AppStarting
                btnTanksPipes.PerformClick()
                ValidateTankPipeTerm(strTankPipeTermErr, bolAllOptionalTankPipeErr)
                Cursor.Current = Cursors.Default
            End If
            If strTankPipeTermErr.Length > 0 And Not bolAllOptionalTankPipeErr Then
                MsgBox(strTankPipeTermErr)
                bolReturnValue = False
            Else
                If strTankPipeTermErr.Length > 0 And bolAllOptionalTankPipeErr Then
                    If MessageBoxCustom.Show(strTankPipeTermErr + vbCrLf + vbCrLf + "Do you want to continue?", "Tank / Pipe Validation", MessageBoxButtons.YesNo, , HorizontalAlignment.Center) = DialogResult.No Then
                        Return False
                    End If
                End If
                Cursor.Current = Cursors.AppStarting
                ValidateResponses(strResponseErr)
                Cursor.Current = Cursors.Default
                If strResponseErr.Length > 0 Then
                    strResponseErr = "The following Line# have no response:" + vbCrLf + strResponseErr.Trim.TrimEnd(",")
                    Dim mbc As MessageBoxCustom
                    If bolAllowYesNo Then
                        Dim dResults As DialogResult = mbc.Show(strResponseErr + vbCrLf + "Do you wish to continue?", "Incomplete Data", MessageBoxButtons.YesNo)
                        If dResults = MsgBoxResult.No Then
                            bolReturnValue = False
                        ElseIf dResults = MsgBoxResult.Yes Then
                            bolReturnValue = True
                        End If
                    Else
                        mbc.Show(strResponseErr + vbCrLf + "You cannot continue", "Incomplete Data", MessageBoxButtons.OK, MsgBoxStyle.Exclamation)
                        bolReturnValue = False
                    End If
                Else
                    bolReturnValue = True
                End If
            End If
            Return bolReturnValue
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ValidateLatLong() As Boolean
        Dim bolFacLatLongSuccess As Boolean = False
        Try
            ' validate lat long
            If txtLatitudeDegree.Text <> String.Empty And _
                txtLatitudeMin.Text <> String.Empty And _
                txtLatitudeSec.Text <> String.Empty And _
                txtLongitudeDegree.Text <> String.Empty And _
                txtLongitudeMin.Text <> String.Empty And _
                txtLongitudeSec.Text <> String.Empty Then
                bolFacLatLongSuccess = True

                If oInspection.CheckListMaster.Owner.Facilities.Datum = 0 Then
                    oInspection.CheckListMaster.Owner.Facilities.Datum = 581
                End If
                If oInspection.CheckListMaster.Owner.Facilities.Method = 0 Then
                    oInspection.CheckListMaster.Owner.Facilities.Method = 583
                End If
                If oInspection.CheckListMaster.Owner.Facilities.LocationType = 0 Then
                    oInspection.CheckListMaster.Owner.Facilities.LocationType = 588
                End If
                If Date.Compare(oInspection.CheckListMaster.Owner.Facilities.DateReceived, dtNullDate) = 0 Then
                    oInspection.CheckListMaster.Owner.Facilities.DateReceived = Now.Date
                End If

            Else
                If txtLatitudeDegree.Text = String.Empty And _
                    txtLatitudeMin.Text = String.Empty And _
                    txtLatitudeSec.Text = String.Empty And _
                    txtLongitudeDegree.Text = String.Empty And _
                    txtLongitudeMin.Text = String.Empty And _
                    txtLongitudeSec.Text = String.Empty Then
                    bolFacLatLongSuccess = True

                    If oInspection.CheckListMaster.Owner.Facilities.Datum <> 0 Then
                        oInspection.CheckListMaster.Owner.Facilities.Datum = 0
                    End If
                    If oInspection.CheckListMaster.Owner.Facilities.Method <> 0 Then
                        oInspection.CheckListMaster.Owner.Facilities.Method = 0
                    End If
                    If oInspection.CheckListMaster.Owner.Facilities.LocationType <> 0 Then
                        oInspection.CheckListMaster.Owner.Facilities.LocationType = 0
                    End If

                Else
                    bolFacLatLongSuccess = False
                    MsgBox("Facility Lat Long validation failed", MsgBoxStyle.OKOnly, "Validation Error")
                End If
            End If
            Return bolFacLatLongSuccess
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub SaveSOCValuesToObject()
        Dim nRowValue As Integer = -1
        For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugSOC.Rows
            nRowValue = IIf(ugRow.Cells("Yes").Value, 1, IIf(ugRow.Cells("No").Value, 0, -1))
            If oInspection.CheckListMaster.InspectionSOC.ID <> ugRow.Cells("ID").Value Then
                oInspection.CheckListMaster.InspectionSOC.Retrieve(oInspection.InspectionInfo, ugRow.Cells("ID").Value)
            End If
            If ugRow.Index = 0 Then
                oInspection.CheckListMaster.InspectionSOC.LeakPrevention = nRowValue
                oInspection.CheckListMaster.InspectionSOC.LeakPreventionCitation = ugRow.Cells("Citations").Text
                oInspection.CheckListMaster.InspectionSOC.LeakPreventionLineNumbers = ugRow.Cells("Line Numbers").Text
            ElseIf ugRow.Index = 1 Then
                oInspection.CheckListMaster.InspectionSOC.LeakDetection = nRowValue
                oInspection.CheckListMaster.InspectionSOC.LeakDetectionCitation = ugRow.Cells("Citations").Text
                oInspection.CheckListMaster.InspectionSOC.LeakDetectionLineNumbers = ugRow.Cells("Line Numbers").Text
            ElseIf ugRow.Index = 2 Then
                oInspection.CheckListMaster.InspectionSOC.LeakPreventionDetection = nRowValue
            End If
        Next
    End Sub
    Private Sub ShowHideMW()
        Try
            btnMW.Visible = oInspection.CheckListMaster.ShowMW
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Navigation Button Click"
    Private Sub CheckTankPipeTermModified()
        Try
            If regenerateCheckListItems Then
                Dim strErr As String = String.Empty
                ValidateTankPipeTerm(strErr)
                If strErr <> String.Empty Then
                    If MessageBoxCustom.Show(strErr + vbCrLf + vbCrLf + "Do you want to Continue?", "Tank / Pipe Cap Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, HorizontalAlignment.Left) = DialogResult.No Then
                        Exit Sub
                    End If
                End If
                Cursor.Current = Cursors.AppStarting
                oInspection.CheckListMaster.GenerateCheckList(True, bolReadOnly)
                ShowHideMW()
                regenerateCheckListItems = False

                ' save tank/pipe/term to archive table only if boolean is set to save from tank/pipe/term grid
                If bolSaveDataToArchiveTbls Then
                    oInspection.PutInspectionArchive(oInspection.FacilityID, oInspection.InspectionInfo.ID, moduleID, MusterContainer.AppUser.UserKey, returnVal, True)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        bolSaveDataToArchiveTbls = False
                        Exit Sub
                    End If
                    bolSaveDataToArchiveTbls = False
                End If

            End If
        Catch ex As Exception
            Throw ex
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub btnMaster_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMaster.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            HideVisiblePanels(True, pnlMaster)
            ToggleButtonAppearance(btnMaster)
            LoadChecklistMaster()
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnTanksPipes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTanksPipes.Click
        Try
            bolFromBtnTanksPipes = True
            Cursor.Current = Cursors.AppStarting
            PrepTankPipeTerm()
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolFromBtnTanksPipes = False
        End Try
    End Sub
    Private Sub btnReg_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReg.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnReg)
            HideVisiblePanels(True, pnlReg)
            Me.ugReg.DataSource = oInspection.CheckListMaster.RegTable(bolReadOnly)
            ugReg.DrawFilter = rp
            SetupGrid(ugReg)
            If ugReg.Rows.Count > 1 Then
                ugReg.Focus()
                ugReg.Rows(1).Cells("Question").Activate()
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub btnCitations_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCitations.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnCitations)
            HideVisiblePanels(True, pnlInspectionCitations)
            Me.ugCitation.DataSource = oInspection.CheckListMaster.CitationTable(bolReadOnly)
            ugCitation.DrawFilter = rp
            SetupGrid(ugCitation)
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCp.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnCp)
            HideVisiblePanels(True, pnlCP)
            pnlCPAdd.Visible = False
            btnAddTankCP.Enabled = False
            btnAddPipeCP.Enabled = False
            btnAddTermCP.Enabled = False
            hasCPTank = False
            hasCPPipe = False
            hasCPTerm = False
            Dim ds As DataSet = oInspection.CheckListMaster.CPTable(bolReadOnly)
            ' #2821 only the cp tested by inspector, show remaining bands
            If ds.Tables("CPTankInspectorTested").Rows.Count > 0 Then
                If ds.Tables("CPTankInspectorTested").Rows(0)("Yes") Then
                    If ds.Tables("CPTank").Rows.Count > 0 Then
                        hasCPTank = True
                    End If
                End If
            End If
            If ds.Tables("CPPipeInspectorTested").Rows.Count > 0 Then
                If ds.Tables("CPPipeInspectorTested").Rows(0)("Yes") Then
                    If ds.Tables("CPPipe").Rows.Count > 0 Then
                        hasCPPipe = True
                    End If
                End If
            End If
            If ds.Tables("CPTermInspectorTested").Rows.Count > 0 Then
                If ds.Tables("CPTermInspectorTested").Rows(0)("Yes") Then
                    If ds.Tables("CPTerm").Rows.Count > 0 Then
                        hasCPTerm = True
                    End If
                End If
            End If
            Me.ugCP.DataSource = ds
            ugCP.DrawFilter = rp
            Me.ugCP.Rows.ExpandAll(True)
            Me.ugCP.DisplayLayout.Override.ExpansionIndicator = Infragistics.Win.UltraWinGrid.ShowExpansionIndicator.Never
            SetupGrid(ugCP)
            ' grid has rectifier row
            ' the foll code is in SetupGrid
            'If ds.Tables("CPRect").Rows.Count > 0 Then
            '    For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCP.Rows
            '        If dr.Cells("Line#").Value = "3.4" Then
            '            If Not dr.ChildBands Is Nothing Then
            '                If Not dr.ChildBands(0).Rows Is Nothing Then
            '                    For Each drChild As Infragistics.Win.UltraWinGrid.UltraGridRow In dr.ChildBands(0).Rows
            '                        SetupRectifierRow(drChild)
            '                    Next
            '                End If
            '            End If
            '            Exit For
            '        End If
            '    Next
            'End If
            If ugCP.Rows.Count > 1 Then
                ugCP.Focus()
                ugCP.Rows(1).Cells("Question").Activate()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSpill_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSpill.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnSpill)
            HideVisiblePanels(True, pnlSpill)
            Me.ugSpill.DataSource = oInspection.CheckListMaster.SpillTable(bolReadOnly)
            ugSpill.DrawFilter = rp
            SetupGrid(ugSpill)
            If ugSpill.Rows.Count > 1 Then
                ugSpill.Focus()
                ugSpill.Rows(1).Cells("Question").Activate()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnTankLeak_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTankLeak.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnTankLeak)
            HideVisiblePanels(True, pnlTankLeak)

            pnlTankLeakAdd.Visible = False
            btnAddTankMW.Enabled = False

            Me.ugTankLeak.DataSource = oInspection.CheckListMaster.TankLeakTable(bolReadOnly)
            ugTankLeak.DrawFilter = rp
            Me.ugTankLeak.Rows.ExpandAll(False)
            Me.ugTankLeak.DisplayLayout.Override.ExpansionIndicator = Infragistics.Win.UltraWinGrid.ShowExpansionIndicator.Never
            SetupGrid(ugTankLeak)
            ' the foll code is in SetupGrid
            'For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ugTankLeak.Rows
            '    If dr.Cells("Line#").Value = "4.2.8" Then
            '        dr.Expanded = True
            '        For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In dr.ChildBands(0).Rows
            '            SetupTankPipeLeakRow(ugChildRow)
            '        Next
            '    Else
            '        dr.Expanded = False
            '    End If
            'Next
            If ugTankLeak.Rows.Count > 1 Then
                ugTankLeak.Focus()
                ugTankLeak.Rows(1).Cells("Question").Activate()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnPipeLeak_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPipeLeak.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnPipeLeak)
            HideVisiblePanels(True, pnlPipeLeak)

            pnlTankLeakAdd.Visible = False
            btnAddTankMW.Enabled = False

            Me.ugPipeLeak.DataSource = oInspection.CheckListMaster.PipeLeakTable(bolReadOnly)
            ugPipeLeak.DrawFilter = rp
            Me.ugPipeLeak.Rows.ExpandAll(False)
            Me.ugPipeLeak.DisplayLayout.Override.ExpansionIndicator = Infragistics.Win.UltraWinGrid.ShowExpansionIndicator.Never
            SetupGrid(ugPipeLeak)
            ' the foll code is in SetupGrid
            'For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ugPipeLeak.Rows
            '    If dr.Cells("Line#").Value = "5.2.8" Then
            '        dr.Expanded = True
            '        For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In dr.ChildBands(0).Rows
            '            SetupTankPipeLeakRow(ugChildRow)
            '        Next
            '    Else
            '        dr.Expanded = False
            '    End If
            'Next
            If ugPipeLeak.Rows.Count > 1 Then
                ugPipeLeak.Focus()
                ugPipeLeak.Rows(1).Cells("Question").Activate()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnVisual_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnVisual.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnVisual)
            HideVisiblePanels(True, pnlVisual)
            Me.ugVisual.DataSource = oInspection.CheckListMaster.VisualTable(bolReadOnly)
            ugVisual.DrawFilter = rp
            SetupGrid(ugVisual)
            If ugVisual.Rows.Count > 1 Then
                ugVisual.Focus()
                ugVisual.Rows(1).Cells("Question").Activate()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnTos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTos.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnTos)
            HideVisiblePanels(True, pnlTOS)
            Me.ugTOS.DataSource = oInspection.CheckListMaster.TOSTable(bolReadOnly)
            ugTOS.DrawFilter = rp
            SetupGrid(ugTOS)
            If ugTOS.Rows.Count > 1 Then
                ugTOS.Focus()
                ugTOS.Rows(1).Cells("Question").Activate()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnComments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnComments.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnComments)
            HideVisiblePanels(True, pnlComments)
            Me.txtComments.Text = oInspection.CheckListMaster.InspectionComments.InsComments
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnMW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMW.Click
        Try
            CheckTankPipeTermModified()
            If btnMW.Visible Then
                Cursor.Current = Cursors.AppStarting
                ToggleButtonAppearance(btnMW)
                HideVisiblePanels(True, pnlMW)
                Me.ugMW.DataSource = oInspection.CheckListMaster.MWellTable(bolReadOnly)
                ugMW.DrawFilter = rp
                SetupGrid(ugMW)
                If ugMW.Rows.Count > 1 Then
                    ugMW.Focus()
                    ugMW.Rows(1).Cells("Question").Activate()
                End If
            End If
        Catch ex As Exception
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub btnSoc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSoc.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnSoc)
            HideVisiblePanels(True, pnlSOC)
            ' only inspection head can edit the values
            ugSOC.DataSource = oInspection.CheckListMaster.SOCTable(Not MusterContainer.AppUser.HEAD_INSPECTION Or bolReadOnly)
            ugSOC.DrawFilter = rp
            SetupGrid(ugSOC)
            If ugSOC.Rows.Count > 1 Then
                ugSOC.Focus()
                ugSOC.Rows(1).Cells("Question").Activate()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCatLeak_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCatLeak.Click
        Try
            CheckTankPipeTermModified()
            Cursor.Current = Cursors.AppStarting
            ToggleButtonAppearance(btnCatLeak)
            HideVisiblePanels(True, pnlCatLeak)
            Me.ugCatLeak.DataSource = oInspection.CheckListMaster.CATLeakTable(bolReadOnly)
            ugCatLeak.DrawFilter = rp
            SetupGrid(ugCatLeak)
            If ugCatLeak.Rows.Count > 1 Then
                ugCatLeak.Focus()
                ugCatLeak.Rows(1).Cells("Question").Activate()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim myErr As New ErrorReport(ex)
            myErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "CheckList"
    Private Sub ugReg_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugReg.CellChange
        Try
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row)
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub ugSpill_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugSpill.CellChange
        Try
            ' 2.6 - show only if response to 2.5 is yes
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row)
                    If e.Cell.Row.Cells("Line#").Value = "2.5" Then
                        oInspection.CheckListMaster.ResponseOfCLItem2point5 = 0
                        If oInspection.CheckListMaster.QuestionIDOfCLItem2point6 <> 0 Then
                            If Not oInspection.InspectionInfo.ChecklistMasterCollection.Item(oInspection.CheckListMaster.QuestionIDOfCLItem2point6) Is Nothing Then
                                Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo = oInspection.InspectionInfo.ChecklistMasterCollection.Item(oInspection.CheckListMaster.QuestionIDOfCLItem2point6)
                                'oInspection.InspectionInfo.ChecklistMasterCollection.Item(oInspection.CheckListMaster.QuestionIDOfCLItem2point6).Show = True
                                'oInspection.InspectionInfo.ChecklistMasterCollection.Item(oInspection.CheckListMaster.QuestionIDOfCLItem2point6).Deleted = False
                                oInspection.CheckListMaster.ShowCLItem(checkList, True, True, bolReadOnly)
                                btnSpill_Click(btnSpill, New System.EventArgs)
                            End If
                        End If
                    End If
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row)
                    If e.Cell.Row.Cells("Line#").Value = "2.5" Then
                        oInspection.CheckListMaster.ResponseOfCLItem2point5 = 0
                        If oInspection.CheckListMaster.QuestionIDOfCLItem2point6 <> 0 Then
                            If Not oInspection.InspectionInfo.ChecklistMasterCollection.Item(oInspection.CheckListMaster.QuestionIDOfCLItem2point6) Is Nothing Then
                                Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo = oInspection.InspectionInfo.ChecklistMasterCollection.Item(oInspection.CheckListMaster.QuestionIDOfCLItem2point6)
                                'oInspection.InspectionInfo.ChecklistMasterCollection.Item(oInspection.CheckListMaster.QuestionIDOfCLItem2point6).Show = False
                                'oInspection.InspectionInfo.ChecklistMasterCollection.Item(oInspection.CheckListMaster.QuestionIDOfCLItem2point6).Deleted = True
                                oInspection.CheckListMaster.ShowCLItem(checkList, False, True, bolReadOnly)
                                btnSpill_Click(btnSpill, New System.EventArgs)
                            End If
                        End If
                    End If
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCP_CellChange(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugCP.CellChange
        Try
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If e.Cell.Band.Index = 0 Then
                    If e.Cell.Row.Cells("RESPONSE").Value = 1 And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        CellUpdate(1, e.Cell.Row)
                    End If
                ElseIf e.Cell.Band.Index = 2 Or e.Cell.Band.Index = 6 Or e.Cell.Band.Index = 10 Then
                    If e.Cell.Row.Cells("TESTED_BY_INSPECTOR_RESPONSE").Value = True And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        CellUpdate(1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
                        If e.Cell.Band.Index = 2 Then
                            hasCPTank = True
                        ElseIf e.Cell.Band.Index = 6 Then
                            hasCPPipe = True
                        ElseIf e.Cell.Band.Index = 10 Then
                            hasCPTerm = True
                        End If
                        SetupGrid(ugCP)
                    End If
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If e.Cell.Band.Index = 0 Then
                    If e.Cell.Row.Cells("RESPONSE").Value = 0 And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        CellUpdate(0, e.Cell.Row)
                    End If
                ElseIf e.Cell.Band.Index = 2 Or e.Cell.Band.Index = 6 Or e.Cell.Band.Index = 10 Then
                    If e.Cell.Row.Cells("TESTED_BY_INSPECTOR_RESPONSE").Value = False And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        CellUpdate(0, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
                        If e.Cell.Band.Index = 2 Then
                            hasCPTank = False
                        ElseIf e.Cell.Band.Index = 6 Then
                            hasCPPipe = False
                        ElseIf e.Cell.Band.Index = 10 Then
                            hasCPTerm = False
                        End If
                        SetupGrid(ugCP)
                    End If
                End If
            ElseIf "Pass".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("PASSFAILINCON").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
                End If
            ElseIf "Fail".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("PASSFAILINCON").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
                End If
            ElseIf "Incon".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("PASSFAILINCON").Value = 2 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(2, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
                End If
            ElseIf "Tank#".Equals(e.Cell.Column.Key) Then
                'If e.Cell.Row.IsAddRow Then
                '    CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
                'Else
                '    e.Cell.CancelUpdate()
                'End If
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
            ElseIf "Pipe#".Equals(e.Cell.Column.Key) Then
                'If e.Cell.Row.IsAddRow Then
                '    CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
                'Else
                '    e.Cell.CancelUpdate()
                'End If
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
            ElseIf "Term#".Equals(e.Cell.Column.Key) Then
                'If e.Cell.Row.IsAddRow Then
                '    CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
                'Else
                '    e.Cell.CancelUpdate()
                'End If
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
            ElseIf "Description of Remote Reference Cell Placement".Equals(e.Cell.Column.Key) Then
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
            ElseIf "Galvanic".Equals(e.Cell.Column.Key) Then
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
            ElseIf "Impressed Current".Equals(e.Cell.Column.Key) Then
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
            ElseIf Not "CCAT".Equals(e.Cell.Column.Key) Then
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugCP.Name)
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTankLeak_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTankLeak.CellChange
        Try
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row)
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row)
                End If
            ElseIf ("Surface Sealed" + vbCrLf + "Yes").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("SURFACE_SEALED").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row, e.Cell.Column.Key, ugTankLeak.Name)
                End If
            ElseIf ("Surface Sealed" + vbCrLf + "No").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("SURFACE_SEALED").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row, e.Cell.Column.Key, ugTankLeak.Name)
                End If
            ElseIf ("Well Caps" + vbCrLf + "Yes").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("WELL_CAPS").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row, e.Cell.Column.Key, ugTankLeak.Name)
                End If
            ElseIf ("Well Caps" + vbCrLf + "No").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("WELL_CAPS").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row, e.Cell.Column.Key, ugTankLeak.Name)
                End If
            ElseIf Not "CCAT".Equals(e.Cell.Column.Key) Then
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugTankLeak.Name)
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugPipeLeak_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugPipeLeak.CellChange
        Try
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row)
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row)
                End If
            ElseIf ("Surface Sealed" + vbCrLf + "Yes").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("SURFACE_SEALED").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row, e.Cell.Column.Key, ugPipeLeak.Name)
                End If
            ElseIf ("Surface Sealed" + vbCrLf + "No").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("SURFACE_SEALED").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row, e.Cell.Column.Key, ugPipeLeak.Name)
                End If
            ElseIf ("Well Caps" + vbCrLf + "Yes").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("WELL_CAPS").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row, e.Cell.Column.Key, ugPipeLeak.Name)
                End If
            ElseIf ("Well Caps" + vbCrLf + "No").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("WELL_CAPS").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row, e.Cell.Column.Key, ugPipeLeak.Name)
                End If
            ElseIf Not "CCAT".Equals(e.Cell.Column.Key) Then
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugPipeLeak.Name)
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCatLeak_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugCatLeak.CellChange
        Try
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row)
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row)
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugVisual_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugVisual.CellChange
        Try
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row)
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row)
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTOS_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTOS.CellChange
        Try
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row)
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("RESPONSE").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row)
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugSOC_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugSOC.CellChange
        'Dim socInfo As MUSTER.Info.InspectionSOCInfo
        Try
            Cursor.Current = Cursors.AppStarting
            If "Yes".Equals(e.Cell.Column.Key) Then
                If oInspection.CheckListMaster.InspectionSOC.ID <> e.Cell.Row.Cells("ID").Value Then
                    oInspection.CheckListMaster.InspectionSOC.Retrieve(oInspection.InspectionInfo, e.Cell.Row.Cells("ID").Value)
                End If
                'socInfo = oInspection.InspectionInfo.SOCsCollection.Item(e.Cell.Row.Cells("ID").Value)
                If e.Cell.Row.Index = 0 Then
                    If oInspection.CheckListMaster.InspectionSOC.LeakPrevention = 1 And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        oInspection.CheckListMaster.InspectionSOC.LeakPrevention = 1
                        oInspection.CheckListMaster.InspectionSOC.CAEOverride = True
                        e.Cell.Value = e.Cell.Text
                        e.Cell.Row.Cells("NO").Value = False
                    End If
                ElseIf e.Cell.Row.Index = 1 Then
                    If oInspection.CheckListMaster.InspectionSOC.LeakDetection = 1 And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        oInspection.CheckListMaster.InspectionSOC.LeakDetection = 1
                        oInspection.CheckListMaster.InspectionSOC.CAEOverride = True
                        e.Cell.Value = e.Cell.Text
                        e.Cell.Row.Cells("NO").Value = False
                    End If
                ElseIf e.Cell.Row.Index = 2 Then
                    If oInspection.CheckListMaster.InspectionSOC.LeakPreventionDetection = 1 And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        oInspection.CheckListMaster.InspectionSOC.LeakPreventionDetection = 1
                        oInspection.CheckListMaster.InspectionSOC.CAEOverride = True
                        e.Cell.Value = e.Cell.Text
                        e.Cell.Row.Cells("NO").Value = False
                    End If
                End If
            ElseIf "No".Equals(e.Cell.Column.Key) Then
                If oInspection.CheckListMaster.InspectionSOC.ID <> e.Cell.Row.Cells("ID").Value Then
                    oInspection.CheckListMaster.InspectionSOC.Retrieve(oInspection.InspectionInfo, e.Cell.Row.Cells("ID").Value)
                End If
                'socInfo = oInspection.InspectionInfo.SOCsCollection.Item(e.Cell.Row.Cells("ID").Value)
                If e.Cell.Row.Index = 0 Then
                    If oInspection.CheckListMaster.InspectionSOC.LeakPrevention = 0 And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        oInspection.CheckListMaster.InspectionSOC.LeakPrevention = 0
                        oInspection.CheckListMaster.InspectionSOC.CAEOverride = True
                        e.Cell.Value = e.Cell.Text
                        e.Cell.Row.Cells("YES").Value = False
                    End If
                ElseIf e.Cell.Row.Index = 1 Then
                    If oInspection.CheckListMaster.InspectionSOC.LeakDetection = 0 And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        oInspection.CheckListMaster.InspectionSOC.LeakDetection = 0
                        oInspection.CheckListMaster.InspectionSOC.CAEOverride = True
                        e.Cell.Value = e.Cell.Text
                        e.Cell.Row.Cells("YES").Value = False
                    End If
                ElseIf e.Cell.Row.Index = 2 Then
                    If oInspection.CheckListMaster.InspectionSOC.LeakPreventionDetection = 0 And e.Cell.Value Then
                        e.Cell.CancelUpdate()
                    Else
                        oInspection.CheckListMaster.InspectionSOC.LeakPreventionDetection = 0
                        oInspection.CheckListMaster.InspectionSOC.CAEOverride = True
                        e.Cell.Value = e.Cell.Text
                        e.Cell.Row.Cells("YES").Value = False
                    End If
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugMW_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugMW.CellChange
        Try
            Cursor.Current = Cursors.AppStarting
            If ("Surface Sealed" + vbCrLf + "Yes").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("SURFACE_SEALED").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row, e.Cell.Column.Key, ugMW.Name)
                End If
            ElseIf ("Surface Sealed" + vbCrLf + "No").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("SURFACE_SEALED").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row, e.Cell.Column.Key, ugMW.Name)
                End If
            ElseIf ("Well Caps" + vbCrLf + "Yes").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("WELL_CAPS").Value = 1 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(1, e.Cell.Row, e.Cell.Column.Key, ugMW.Name)
                End If
            ElseIf ("Well Caps" + vbCrLf + "No").Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("WELL_CAPS").Value = 0 And e.Cell.Value Then
                    e.Cell.CancelUpdate()
                Else
                    CellUpdate(0, e.Cell.Row, e.Cell.Column.Key, ugMW.Name)
                End If
            ElseIf Not "CCAT".Equals(e.Cell.Column.Key) Then
                CellUpdate(-1, e.Cell.Row, e.Cell.Column.Key, ugMW.Name)
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub CellUpdate(ByVal response As Int64, ByRef row As Infragistics.Win.UltraWinGrid.UltraGridRow, Optional ByVal ugCell As String = "", Optional ByVal gridName As String = "")
        Dim deleteCCat As Boolean = False
        Dim deleteDiscrep As Boolean = False
        Dim deleteCitation As Boolean = False
        Dim createDiscrep As Boolean = True
        Dim createCitation As Boolean = True

        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
        Dim ccat As MUSTER.Info.InspectionCCATInfo
        Dim discrep As MUSTER.Info.InspectionDiscrepInfo
        Dim citation As MUSTER.Info.InspectionCitationInfo
        ' response : 1 = Yes, 0 = No, 2 = Incon, -1 = no checkbox involved in the update
        Try
            checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(row.Cells("QUESTION_ID").Value)

            If response = 1 Then
                ' ccat exists only if checklist's ccat is true
                If checkList.CCAT Then
                    ' check if ccat exists
                    For Each ccat In oInspection.InspectionInfo.CCATsCollection.Values
                        If ccat.QuestionID = checkList.ID Then
                            If Not ccat.Deleted Then
                                If ccat.ID > 0 Or (ccat.ID <= 0 And ccat.IsDirty) Then
                                    deleteCCat = True
                                    Exit For
                                End If
                            End If
                            'Exit For
                        End If
                    Next
                End If

                If ugCell = "" And gridName = "" Then
                    row.Cells("RESPONSE").Value = 1
                    row.Cells("No").Value = False
                    row.Cells("Yes").Value = True
                ElseIf gridName = ugCP.Name Then
                    If ugCell = "Yes" Or ugCell = "No" Then
                        row.Cells("TESTED_BY_INSPECTOR_RESPONSE").Value = True
                        row.Cells("Yes").Value = True
                        row.Cells("No").Value = False
                        Select Case row.ParentRow.Cells("Line#").Text
                            Case "3.5.4"
                                ugCP.DisplayLayout.Bands(3).Hidden = False
                                ugCP.DisplayLayout.Bands(4).Hidden = False
                                ugCP.DisplayLayout.Bands(5).Hidden = False
                                bol354Hidden = False
                            Case "3.6.3"
                                ugCP.DisplayLayout.Bands(7).Hidden = False
                                ugCP.DisplayLayout.Bands(8).Hidden = False
                                ugCP.DisplayLayout.Bands(9).Hidden = False
                                bol363Hidden = False
                            Case "3.7.6"
                                ugCP.DisplayLayout.Bands(11).Hidden = False
                                ugCP.DisplayLayout.Bands(12).Hidden = False
                                ugCP.DisplayLayout.Bands(13).Hidden = False
                                bol376Hidden = False
                        End Select
                    Else
                        row.Cells("PASSFAILINCON").Value = 1
                        row.Cells("Pass").Value = True
                        row.Cells("Fail").Value = False
                        row.Cells("Incon").Value = False
                    End If
                ElseIf gridName = ugPipeLeak.Name Or gridName = ugTankLeak.Name Or gridName = ugMW.Name Then
                    Select Case ugCell
                        Case ("Surface Sealed" + vbCrLf + "Yes")
                            row.Cells("SURFACE_SEALED").Value = 1
                            row.Cells("Surface Sealed" + vbCrLf + "Yes").Value = True
                            row.Cells("Surface Sealed" + vbCrLf + "No").Value = False
                        Case ("Well Caps" + vbCrLf + "Yes")
                            row.Cells("WELL_CAPS").Value = 1
                            row.Cells("Well Caps" + vbCrLf + "Yes").Value = True
                            row.Cells("Well Caps" + vbCrLf + "No").Value = False
                    End Select
                End If

                ' if ccat exists for the checklist item
                If deleteCCat Then
                    Dim result As MsgBoxResult = MsgBox("CCAT Exists for Line# " + row.Cells("Line#").Value.ToString + vbCrLf + _
                                                        "If you proceed, it will delete" + vbCrLf + "existing information" + vbCrLf + _
                                                        "Do you want to Continue?", MsgBoxStyle.YesNo)
                    If result = MsgBoxResult.Yes Then
                        ' delete ccat(s)
                        Dim delIDs As New Collection
                        Dim index As Integer
                        For Each ccat In oInspection.InspectionInfo.CCATsCollection.Values
                            If ccat.QuestionID = checkList.ID Then
                                If oInspection.CheckListMaster.InspectionCCAT.ID <> ccat.ID Then
                                    oInspection.CheckListMaster.InspectionCCAT.Retrieve(oInspection.InspectionInfo, ccat.ID)
                                End If
                                oInspection.CheckListMaster.InspectionCCAT.Reset()
                                oInspection.CheckListMaster.InspectionCCAT.Deleted = True
                            End If
                        Next
                    ElseIf result = MsgBoxResult.No Then
                        If ugCell = "" And gridName = "" Then
                            row.Cells("RESPONSE").Value = row.Cells("RESPONSE").OriginalValue
                            row.Cells("No").Value = row.Cells("No").OriginalValue
                            row.Cells("Yes").Value = row.Cells("Yes").OriginalValue
                        ElseIf gridName = "ugCP" Then
                            row.Cells("PASSFAILINCON").Value = row.Cells("PASSFAILINCON").OriginalValue
                            row.Cells("Pass").Value = row.Cells("Pass").OriginalValue
                            row.Cells("Fail").Value = IIf(row.Cells("PASSFAILINCON").Value = 0, True, False)
                            row.Cells("Incon").Value = IIf(row.Cells("PASSFAILINCON").Value = 2, True, False)
                        ElseIf gridName = "ugPipeLeak" Or gridName = "ugTankLeak" Then
                            Select Case ugCell
                                Case ("Surface Sealed" + vbCrLf + "Yes")
                                    row.Cells("SURFACE_SEALED").Value = row.Cells("SURFACE_SEALED").OriginalValue
                                    row.Cells("Surface Sealed" + vbCrLf + "Yes").Value = row.Cells("Surface Sealed" + vbCrLf + "Yes").OriginalValue
                                    row.Cells("Surface Sealed" + vbCrLf + "No").Value = row.Cells("Surface Sealed" + vbCrLf + "No").OriginalValue
                                Case ("Well Caps" + vbCrLf + "Yes")
                                    row.Cells("WELL_CAPS").Value = row.Cells("WELL_CAPS").OriginalValue
                                    row.Cells("Well Caps" + vbCrLf + "Yes").Value = row.Cells("Well Caps" + vbCrLf + "Yes").OriginalValue
                                    row.Cells("Well Caps" + vbCrLf + "No").Value = row.Cells("Well Caps" + vbCrLf + "No").OriginalValue
                            End Select
                        End If
                        Exit Sub
                    End If
                End If

                ' if there are any citations / discreps - delete
                ' citation exists only if citation is not equal to -1
                ' discrep exists only if DiscrepText is not empty
                Dim strUpdatedCCAT As String = ""
                If checkList.Citation <> -1 Or checkList.DiscrepText <> String.Empty Then
                    Dim slCCAT As New SortedList
                    If ugCell = "" And gridName = "" Then
                        ' citation and discrep are unique for each row inside this condition
                        deleteCitation = True
                        deleteDiscrep = True
                    ElseIf gridName = ugCP.Name Then
                        deleteCitation = True
                        deleteDiscrep = True

                        For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In row.ParentCollection
                            ' current row is not the same row being updated
                            If ugRow.Cells("ID").Value <> row.Cells("ID").Value Then
                                ' if any other row belonging to the updating row's parent
                                ' has "Fail" or "Incon" set to true, set delete variables to false
                                If ugRow.Cells("Fail").Value Or ugRow.Cells("Incon").Value Then
                                    If ugRow.Cells("Line#").Text.StartsWith("3.5.4.") Then
                                        'strUpdatedCCAT += "T" + ugRow.Cells("Tank#").Text + ", "
                                        If Not slCCAT.Contains(ugRow.Cells("Tank#").Value) Then
                                            slCCAT.Add(ugRow.Cells("Tank#").Value, "T" + ugRow.Cells("Tank#").Text)
                                        End If
                                    ElseIf ugRow.Cells("Line#").Text.StartsWith("3.6.3.") Then
                                        'strUpdatedCCAT += "P" + ugRow.Cells("Pipe#").Text + ", "
                                        If Not slCCAT.Contains(ugRow.Cells("Pipe#").Value) Then
                                            slCCAT.Add(ugRow.Cells("Pipe#").Value, "P" + ugRow.Cells("Pipe#").Text)
                                        End If
                                    ElseIf ugRow.Cells("Line#").Text.StartsWith("3.7.6.") Then
                                        'strUpdatedCCAT += "TP" + ugRow.Cells("Term#").Text + ", "
                                        If Not slCCAT.Contains(ugRow.Cells("Term#").Value) Then
                                            slCCAT.Add(ugRow.Cells("Term#").Value, "TP" + ugRow.Cells("Term#").Text)
                                        End If
                                    Else
                                        deleteCitation = False
                                    End If
                                    deleteDiscrep = False
                                End If
                            End If
                        Next

                        For i As Integer = 0 To slCCAT.Count - 1
                            strUpdatedCCAT += slCCAT.GetByIndex(i).ToString + ", "
                        Next
                        If strUpdatedCCAT <> String.Empty Then
                            strUpdatedCCAT = strUpdatedCCAT.Trim.TrimEnd(",")
                        End If

                    ElseIf gridName = ugPipeLeak.Name Or gridName = ugTankLeak.Name Then
                        deleteCitation = True
                        deleteDiscrep = True
                        For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In row.ParentCollection
                            ' current row is not the same row being updated
                            If ugRow.Cells("ID").Value <> row.Cells("ID").Value Then
                                ' if any other row belonging to the updating row's parent
                                ' has "No" set to true, set delete variables to false
                                If ugRow.Cells("Surface Sealed" + vbCrLf + "No").Value Or _
                                    ugRow.Cells("Well Caps" + vbCrLf + "No").Value Then

                                    If Not slCCAT.Contains(ugRow.Cells("Well#").Value) Then
                                        slCCAT.Add(ugRow.Cells("Well#").Value, ugRow.Cells("Well#").Text)
                                    End If

                                    'deleteCitation = False
                                    deleteDiscrep = False
                                End If
                            End If
                        Next

                        For i As Integer = 0 To slCCAT.Count - 1
                            strUpdatedCCAT += slCCAT.GetByIndex(i).ToString + ", "
                        Next
                        If strUpdatedCCAT <> String.Empty Then
                            strUpdatedCCAT = strUpdatedCCAT.Trim.TrimEnd(",")
                        End If

                    End If
                End If

                If deleteCitation Then
                    For Each citation In oInspection.InspectionInfo.CitationsCollection.Values
                        If citation.QuestionID = checkList.ID Then
                            If oInspection.CheckListMaster.InspectionCitation.ID <> citation.ID Then
                                oInspection.CheckListMaster.InspectionCitation.Retrieve(oInspection.InspectionInfo, citation.ID)
                            End If
                            If strUpdatedCCAT <> String.Empty Then
                                'citation.Reset()
                                oInspection.CheckListMaster.InspectionCitation.CCAT = strUpdatedCCAT
                                oInspection.CheckListMaster.InspectionCitation.Deleted = False
                            Else
                                oInspection.CheckListMaster.InspectionCitation.Reset()
                                oInspection.CheckListMaster.InspectionCitation.CCAT = String.Empty
                                oInspection.CheckListMaster.InspectionCitation.Deleted = True
                            End If
                            Exit For
                        End If
                    Next
                    'If Not (citation Is Nothing) Then
                    '    If citation.ID < 0 Then
                    '        oInspection.CheckListMaster.InspectionCitation.Remove(citation)
                    '    End If
                    'End If
                End If

                If deleteDiscrep Then
                    For Each discrep In oInspection.InspectionInfo.DiscrepsCollection.Values
                        If discrep.QuestionID = checkList.ID Then
                            If oInspection.CheckListMaster.InspectionDiscrep.ID <> discrep.ID Then
                                oInspection.CheckListMaster.InspectionDiscrep.Retrieve(oInspection.InspectionInfo, discrep.ID)
                            End If
                            oInspection.CheckListMaster.InspectionDiscrep.Reset()
                            oInspection.CheckListMaster.InspectionDiscrep.Deleted = True
                            Exit For
                        End If
                    Next
                    'If Not (discrep Is Nothing) Then
                    '    If discrep.ID < 0 Then
                    '        oInspection.CheckListMaster.InspectionDiscrep.Remove(discrep)
                    '    End If
                    'End If
                End If

                ' Assign value to Object
                If ugCell = "" And gridName = "" Then
                    If oInspection.CheckListMaster.InspectionResponses.ID <> CType(row.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionResponses.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                    End If
                    oInspection.CheckListMaster.InspectionResponses.Response = row.Cells("RESPONSE").Value
                ElseIf gridName = ugCP.Name Then
                    If oInspection.CheckListMaster.InspectionCPReadings.ID <> CType(row.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                    End If
                    If ugCell = "Yes" Or ugCell = "No" Then
                        oInspection.CheckListMaster.InspectionCPReadings.TestedByInspectorResponse = row.Cells("TESTED_BY_INSPECTOR_RESPONSE").Value
                        oInspection.CheckListMaster.AddCPReadings(True, bolReadOnly)
                    Else
                        oInspection.CheckListMaster.InspectionCPReadings.PassFailIncon = row.Cells("PASSFAILINCON").Value
                    End If
                ElseIf gridName = ugPipeLeak.Name Or gridName = ugTankLeak.Name Or gridName = ugMW.Name Then
                    If oInspection.CheckListMaster.InspectionMonitorWells.ID <> CType(row.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionMonitorWells.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                    End If
                    Select Case ugCell
                        Case ("Surface Sealed" + vbCrLf + "Yes")
                            oInspection.CheckListMaster.InspectionMonitorWells.SurfaceSealed = row.Cells("SURFACE_SEALED").Value
                        Case ("Well Caps" + vbCrLf + "Yes")
                            oInspection.CheckListMaster.InspectionMonitorWells.WellCaps = row.Cells("WELL_CAPS").Value
                    End Select
                End If

                If ugCell = "" And gridName = "" Then
                    row.Cells("CCAT").Value = String.Empty
                End If

            ElseIf response = 0 Or response = 2 Then
                If ugCell = "" And gridName = "" Then
                    row.Cells("RESPONSE").Value = 0
                    row.Cells("Yes").Value = False
                    row.Cells("No").Value = True
                ElseIf gridName = ugCP.Name Then
                    If ugCell = "Yes" Or ugCell = "No" Then
                        row.Cells("TESTED_BY_INSPECTOR_RESPONSE").Value = False
                        row.Cells("Yes").Value = False
                        row.Cells("No").Value = True
                        Select Case row.ParentRow.Cells("Line#").Text
                            Case "3.5.4"
                                ugCP.DisplayLayout.Bands(3).Hidden = True
                                ugCP.DisplayLayout.Bands(4).Hidden = True
                                ugCP.DisplayLayout.Bands(5).Hidden = True
                                bol354Hidden = True
                            Case "3.6.3"
                                ugCP.DisplayLayout.Bands(7).Hidden = True
                                ugCP.DisplayLayout.Bands(8).Hidden = True
                                ugCP.DisplayLayout.Bands(9).Hidden = True
                                bol363Hidden = True
                            Case "3.7.6"
                                ugCP.DisplayLayout.Bands(11).Hidden = True
                                ugCP.DisplayLayout.Bands(12).Hidden = True
                                ugCP.DisplayLayout.Bands(13).Hidden = True
                                bol376Hidden = True
                        End Select
                    Else
                        row.Cells("PASSFAILINCON").Value = response
                        row.Cells("Pass").Value = False
                        row.Cells("Fail").Value = IIf(response = 0, True, False)
                        row.Cells("Incon").Value = IIf(response = 2, True, False)
                    End If
                ElseIf gridName = ugPipeLeak.Name Or gridName = ugTankLeak.Name Or gridName = ugMW.Name Then
                    Select Case ugCell
                        Case ("Surface Sealed" + vbCrLf + "No")
                            row.Cells("SURFACE_SEALED").Value = 0
                            row.Cells("Surface Sealed" + vbCrLf + "Yes").Value = False
                            row.Cells("Surface Sealed" + vbCrLf + "No").Value = True
                        Case ("Well Caps" + vbCrLf + "No")
                            row.Cells("WELL_CAPS").Value = 0
                            row.Cells("Well Caps" + vbCrLf + "Yes").Value = False
                            row.Cells("Well Caps" + vbCrLf + "No").Value = True
                    End Select
                End If

                ' Assign value to Object
                If ugCell = "" And gridName = "" Then
                    If oInspection.CheckListMaster.InspectionResponses.ID <> CType(row.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionResponses.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                    End If
                    oInspection.CheckListMaster.InspectionResponses.Response = row.Cells("RESPONSE").Value
                ElseIf gridName = ugCP.Name Then
                    If oInspection.CheckListMaster.InspectionCPReadings.ID <> CType(row.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                    End If
                    If ugCell = "Yes" Or ugCell = "No" Then
                        oInspection.CheckListMaster.InspectionCPReadings.TestedByInspectorResponse = row.Cells("TESTED_BY_INSPECTOR_RESPONSE").Value
                        oInspection.CheckListMaster.AddCPReadings(True, bolReadOnly)
                    Else
                        oInspection.CheckListMaster.InspectionCPReadings.PassFailIncon = response
                    End If
                ElseIf gridName = ugPipeLeak.Name Or gridName = ugTankLeak.Name Or gridName = ugMW.Name Then
                    If oInspection.CheckListMaster.InspectionMonitorWells.ID <> CType(row.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionMonitorWells.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                    End If
                    Select Case ugCell
                        Case ("Surface Sealed" + vbCrLf + "No")
                            oInspection.CheckListMaster.InspectionMonitorWells.SurfaceSealed = row.Cells("SURFACE_SEALED").Value
                        Case ("Well Caps" + vbCrLf + "No")
                            oInspection.CheckListMaster.InspectionMonitorWells.WellCaps = row.Cells("WELL_CAPS").Value
                    End Select
                End If

                ' citation exists only if citation is not equal to -1
                If checkList.Citation <> -1 And Not (ugCell = "Yes" Or ugCell = "No") Then
                    ' citation
                    ' check if citation exists in collection, else create and add to collection
                    For Each citation In oInspection.InspectionInfo.CitationsCollection.Values
                        If citation.QuestionID = checkList.ID Then
                            If oInspection.CheckListMaster.InspectionCitation.ID <> citation.ID Then
                                oInspection.CheckListMaster.InspectionCitation.Retrieve(oInspection.InspectionInfo, citation.ID)
                            End If
                            createCitation = False
                            oInspection.CheckListMaster.InspectionCitation.Deleted = False
                            If row.Cells("Line#").Text.StartsWith("3.5.4.") Or row.Cells("Line#").Text.StartsWith("3.6.3.") Or row.Cells("Line#").Text.StartsWith("3.7.6.") Then
                                If oInspection.CheckListMaster.InspectionCitation.CCAT = String.Empty Then
                                    If row.Cells("Line#").Text.StartsWith("3.5.4.") Then
                                        oInspection.CheckListMaster.InspectionCitation.CCAT = "T" + row.Cells("Tank#").Text
                                    ElseIf row.Cells("Line#").Text.StartsWith("3.6.3.") Then
                                        oInspection.CheckListMaster.InspectionCitation.CCAT = "P" + row.Cells("Pipe#").Text
                                    Else
                                        oInspection.CheckListMaster.InspectionCitation.CCAT = "TP" + row.Cells("Term#").Text
                                    End If
                                Else
                                    If row.Cells("Line#").Text.StartsWith("3.5.4.") Then
                                        oInspection.CheckListMaster.InspectionCitation.CCAT += ", T" + row.Cells("Tank#").Text
                                    ElseIf row.Cells("Line#").Text.StartsWith("3.6.3.") Then
                                        oInspection.CheckListMaster.InspectionCitation.CCAT += ", P" + row.Cells("Pipe#").Text
                                    Else
                                        oInspection.CheckListMaster.InspectionCitation.CCAT += ", TP" + row.Cells("Term#").Text
                                    End If

                                    Dim slCCAT As New SortedList
                                    Dim str As String = String.Empty
                                    For Each str In oInspection.CheckListMaster.InspectionCitation.CCAT.Split(",")
                                        If row.Cells("Line#").Text.StartsWith("3.5.4.") Then
                                            If Not slCCAT.contains(CType(str.Trim.TrimStart("T"), Integer)) Then
                                                slCCAT.Add(CType(str.Trim.TrimStart("T"), Integer), str)
                                            End If
                                        ElseIf row.Cells("Line#").Text.StartsWith("3.6.3.") Then
                                            If Not slCCAT.contains(CType(str.Trim.TrimStart("P"), Integer)) Then
                                                slCCAT.Add(CType(str.Trim.TrimStart("P"), Integer), str)
                                            End If
                                        Else
                                            If Not slCCAT.contains(CType((str.Trim.TrimStart("T")).TrimStart("P"), Integer)) Then
                                                slCCAT.Add(CType((str.Trim.TrimStart("T")).TrimStart("P"), Integer), str)
                                            End If
                                        End If
                                    Next
                                    str = String.Empty
                                    For i As Integer = 0 To slCCAT.Count - 1
                                        str += slCCAT.GetByIndex(i).ToString + ", "
                                    Next
                                    If str <> String.Empty Then
                                        str = str.Trim.TrimEnd(",")
                                    End If
                                    oInspection.CheckListMaster.InspectionCitation.CCAT = str
                                End If
                            ElseIf row.Cells("Line#").Text.StartsWith("4.2.8.") Or row.Cells("Line#").Text.StartsWith("5.2.8.") Then
                                If oInspection.CheckListMaster.InspectionCitation.CCAT = String.Empty Then
                                    oInspection.CheckListMaster.InspectionCitation.CCAT = row.Cells("Well#").Text
                                Else
                                    oInspection.CheckListMaster.InspectionCitation.CCAT += ", " + row.Cells("Well#").Text
                                    Dim slCCAT As New SortedList
                                    Dim str As String = String.Empty
                                    For Each str In oInspection.CheckListMaster.InspectionCitation.CCAT.Split(",")
                                        If Not slCCAT.contains(CType(str, Integer)) Then
                                            slCCAT.Add(CType(str, Integer), str)
                                        End If
                                    Next
                                    str = String.Empty
                                    For i As Integer = 0 To slCCAT.Count - 1
                                        str += slCCAT.GetByIndex(i).ToString + ", "
                                    Next
                                    If str <> String.Empty Then
                                        str = str.Trim.TrimEnd(",")
                                    End If
                                    oInspection.CheckListMaster.InspectionCitation.CCAT = str
                                End If
                            Else
                                oInspection.CheckListMaster.InspectionCitation.Reset()
                            End If
                            Exit For
                        End If
                    Next
                    If createCitation Then
                        citation = New MUSTER.Info.InspectionCitationInfo(0, _
                        oInspection.ID, _
                        checkList.ID, _
                        oInspection.FacilityID, _
                        0, _
                        0, _
                        checkList.Citation, _
                        String.Empty, _
                        False, _
                        dtNullDate, _
                        dtNullDate, _
                        dtNullDate, _
                        False, _
                        String.Empty, _
                        dtNullDate, _
                        String.Empty, _
                        dtNullDate)
                        If row.Cells("Line#").Text.StartsWith("3.5.4.") Then
                            citation.CCAT = "T" + row.Cells("Tank#").Text
                        ElseIf row.Cells("Line#").Text.StartsWith("3.6.3.") Then
                            citation.CCAT = "P" + row.Cells("Pipe#").Text
                        ElseIf row.Cells("Line#").Text.StartsWith("3.7.6.") Then
                            citation.CCAT = "TP" + row.Cells("Term#").Text
                        ElseIf row.Cells("Line#").Text.StartsWith("4.2.8.") Or row.Cells("Line#").Text.StartsWith("5.2.8.") Then
                            citation.CCAT = row.Cells("Well#").Text
                        End If
                        oInspection.CheckListMaster.InspectionCitation.Add(citation)
                    End If
                End If

                ' discrep exists only if DiscrepText is not empty
                If checkList.DiscrepText <> String.Empty And Not (ugCell = "Yes" Or ugCell = "No") Then
                    ' discrep
                    ' check if discrep exists in collection, else create and add to collection
                    For Each discrep In oInspection.InspectionInfo.DiscrepsCollection.Values
                        If discrep.QuestionID = checkList.ID Then
                            If oInspection.CheckListMaster.InspectionDiscrep.ID <> discrep.ID Then
                                oInspection.CheckListMaster.InspectionDiscrep.Retrieve(oInspection.InspectionInfo, discrep.ID)
                            End If
                            createDiscrep = False
                            oInspection.CheckListMaster.InspectionDiscrep.Reset()
                            oInspection.CheckListMaster.InspectionDiscrep.Deleted = False
                            Exit For
                        End If
                    Next
                    If createDiscrep Then
                        discrep = New MUSTER.Info.InspectionDiscrepInfo(0, _
                        oInspection.ID, _
                        checkList.ID, _
                        checkList.DiscrepText, _
                        False, _
                        String.Empty, _
                        dtNullDate, _
                        String.Empty, _
                        dtNullDate, _
                        False, _
                        dtNullDate, _
                        oInspection.CheckListMaster.InspectionCitation.ID)
                        oInspection.CheckListMaster.InspectionDiscrep.Add(discrep)
                    End If
                End If

                ' CCAT
                If checkList.CCAT And Not (ugCell = "Yes" Or ugCell = "No") Then
                    ' reset ccat
                    For Each ccat In oInspection.InspectionInfo.CCATsCollection.Values
                        If ccat.QuestionID = checkList.ID Then
                            If oInspection.CheckListMaster.InspectionCCAT.ID <> ccat.ID Then
                                oInspection.CheckListMaster.InspectionCCAT.Retrieve(oInspection.InspectionInfo, ccat.ID)
                            End If
                            oInspection.CheckListMaster.InspectionCCAT.Reset()
                            oInspection.CheckListMaster.InspectionCCAT.TankPipeResponse = False
                            oInspection.CheckListMaster.InspectionCCAT.TankPipeResponseDetail = String.Empty
                            oInspection.CheckListMaster.InspectionCCAT.Deleted = False
                        End If
                    Next
                    Dim frmCCAT As CCAT
                    frmCCAT = New CCAT(oInspection, row, bolReadOnly, gridName, ugCell)
                    frmCCAT.ShowDialog()
                End If

            ElseIf response = -1 Then
                'Dim cpReadingInfo As MUSTER.Info.InspectionCPReadingsInfo
                'Dim rectifierInfo As MUSTER.Info.InspectionRectifierInfo
                'Dim mwInfo As MUSTER.Info.InspectionMonitorWellsInfo
                Dim bolCancel As Boolean = False
                Select Case gridName
                    Case ugCP.Name
                        Select Case ugCell
                            Case "Volts", "Amps", "Hours", "How Long"
                                If oInspection.CheckListMaster.InspectionRectifier.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionRectifier.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'rectifierInfo = oInspection.InspectionInfo.RectifiersCollection.Item(row.Cells("ID").Value.ToString)
                                Select Case ugCell
                                    Case "Volts"
                                        If IsNumeric(row.Cells("Volts").EditorResolved.Value) Then
                                            oInspection.CheckListMaster.InspectionRectifier.Volts = row.Cells("Volts").EditorResolved.Value
                                        Else
                                            oInspection.CheckListMaster.InspectionRectifier.Volts = 0.0
                                        End If
                                    Case "Amps"
                                        If IsNumeric(row.Cells("Amps").EditorResolved.Value) Then
                                            oInspection.CheckListMaster.InspectionRectifier.Amps = row.Cells("Amps").EditorResolved.Value
                                        Else
                                            oInspection.CheckListMaster.InspectionRectifier.Amps = 0.0
                                        End If
                                    Case "Hours"
                                        If IsNumeric(row.Cells("Hours").EditorResolved.Value) Then
                                            oInspection.CheckListMaster.InspectionRectifier.Hours = row.Cells("Hours").EditorResolved.Value
                                        Else
                                            oInspection.CheckListMaster.InspectionRectifier.Hours = 0.0
                                        End If
                                    Case "How Long"
                                        oInspection.CheckListMaster.InspectionRectifier.InopHowLong = row.Cells("How Long").Text
                                End Select
                                row.Cells("Volts").Value = oInspection.CheckListMaster.InspectionRectifier.Volts
                                row.Cells("Amps").Value = oInspection.CheckListMaster.InspectionRectifier.Amps
                                row.Cells("Hours").Value = oInspection.CheckListMaster.InspectionRectifier.Hours
                                row.Cells("How Long").Value = oInspection.CheckListMaster.InspectionRectifier.InopHowLong
                            Case "Description of Remote Reference Cell Placement"
                                If oInspection.CheckListMaster.InspectionCPReadings.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'cpReadingInfo = oInspection.InspectionInfo.CPReadingsCollection.Item(row.Cells("ID").Value.ToString)
                                oInspection.CheckListMaster.InspectionCPReadings.LocalReferCellPlacement = row.Cells("Description of Remote Reference Cell Placement").Text
                            Case "Galvanic"
                                If oInspection.CheckListMaster.InspectionCPReadings.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'cpReadingInfo = oInspection.InspectionInfo.CPReadingsCollection.Item(row.Cells("ID").Value.ToString)
                                If row.Cells("GALVANIC_IC_RESPONSE").Value = 0 And row.Cells("Galvanic").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.GalvanicICResponse = -1
                                    row.Cells("Galvanic").Value = False
                                    row.Cells("Impressed Current").Value = False
                                    row.Cells("GALVANIC_IC_RESPONSE").Value = -1
                                Else
                                    oInspection.CheckListMaster.InspectionCPReadings.GalvanicICResponse = 0
                                    row.Cells("Galvanic").Value = True
                                    row.Cells("Impressed Current").Value = False
                                    row.Cells("GALVANIC_IC_RESPONSE").Value = 0
                                End If
                            Case "Impressed Current"
                                If oInspection.CheckListMaster.InspectionCPReadings.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'cpReadingInfo = oInspection.InspectionInfo.CPReadingsCollection.Item(row.Cells("ID").Value.ToString)
                                If row.Cells("GALVANIC_IC_RESPONSE").Value = 1 And row.Cells("Impressed Current").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.GalvanicICResponse = -1
                                    row.Cells("Galvanic").Value = False
                                    row.Cells("Impressed Current").Value = False
                                    row.Cells("GALVANIC_IC_RESPONSE").Value = -1
                                Else
                                    oInspection.CheckListMaster.InspectionCPReadings.GalvanicICResponse = 1
                                    row.Cells("Galvanic").Value = False
                                    row.Cells("Impressed Current").Value = True
                                    row.Cells("GALVANIC_IC_RESPONSE").Value = 1
                                End If
                            Case "Contact Point", "Local Reference Cell Placement", "Local/On", "Remote/Off"
                                If oInspection.CheckListMaster.InspectionCPReadings.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'cpReadingInfo = oInspection.InspectionInfo.CPReadingsCollection.Item(row.Cells("ID").Value.ToString)
                                Select Case row.ParentRow.Cells("Line#").Text
                                    Case "3.5.4"
                                        If row.Cells("Tank#").Text = String.Empty Or row.Cells("Tank#").Text Is DBNull.Value Then
                                            MsgBox("Please select Tank first")
                                            bolCancel = True
                                            Exit Select
                                        End If
                                    Case "3.6.3"
                                        If row.Cells("Pipe#").Text = String.Empty Or row.Cells("Pipe#").Text Is DBNull.Value Then
                                            MsgBox("Please select Pipe first")
                                            bolCancel = True
                                            Exit Select
                                        End If
                                    Case "3.7.6"
                                        If row.Cells("Term#").Text = String.Empty Or row.Cells("Term#").Text Is DBNull.Value Then
                                            MsgBox("Please select Term first")
                                            bolCancel = True
                                            Exit Select
                                        End If
                                End Select
                                If bolCancel Then
                                    If row.Cells("Contact Point").Activated Then
                                        row.Cells("Contact Point").CancelUpdate()
                                    ElseIf row.Cells("Local Reference Cell Placement").Activated Then
                                        row.Cells("Local Reference Cell Placement").CancelUpdate()
                                    ElseIf row.Cells("Local/On").Activated Then
                                        row.Cells("Local/On").CancelUpdate()
                                    ElseIf row.Cells("Remote/Off").Activated Then
                                        row.Cells("Remote/Off").CancelUpdate()
                                    End If
                                Else
                                    oInspection.CheckListMaster.InspectionCPReadings.ContactPoint = row.Cells("Contact Point").Text
                                    oInspection.CheckListMaster.InspectionCPReadings.LocalReferCellPlacement = row.Cells("Local Reference Cell Placement").Text
                                    oInspection.CheckListMaster.InspectionCPReadings.LocalOn = row.Cells("Local/On").Text
                                    oInspection.CheckListMaster.InspectionCPReadings.RemoteOff = row.Cells("Remote/Off").Text
                                End If
                            Case "Tank#"
                                If oInspection.CheckListMaster.InspectionCPReadings.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'cpReadingInfo = oInspection.InspectionInfo.CPReadingsCollection.Item(row.Cells("ID").Value.ToString)
                                If row.IsAddRow Then
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeID = row.Cells("Tank#").ValueListResolved.GetValue(row.Cells("Tank#").ValueListResolved.SelectedItemIndex)
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex = row.Cells("Tank#").ValueListResolved.GetText(row.Cells("Tank#").ValueListResolved.SelectedItemIndex)
                                    row.Cells("TANK_INDEX").Value = oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex
                                    ' archive so when you reset, the tank / pipe / term id/index is not lost
                                    ' there is no other data entered before tank / pipe / term is selected
                                    oInspection.CheckListMaster.InspectionCPReadings.Archive()
                                    ' check citation
                                Else
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeID = row.Cells("Tank#").ValueListResolved.GetValue(row.Cells("Tank#").ValueListResolved.SelectedItemIndex)
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex = row.Cells("Tank#").ValueListResolved.GetText(row.Cells("Tank#").ValueListResolved.SelectedItemIndex)
                                    row.Cells("TANK_INDEX").Value = oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex
                                End If
                                ' need to update fuel type with respect to Tank#
                                row.Cells("Fuel Type").Value = oInspection.CheckListMaster.TankFuelType.Item(oInspection.CheckListMaster.InspectionCPReadings.TankPipeID)
                                ' need to update the citaion with correct Tank#
                                UpdateCitationCCATForCP(row, checkList)

                                'If cpReadingInfo.TankPipeID <> 0 And Not row.IsAddRow Then
                                '    Dim result As MsgBoxResult = MsgBox("Do you want to continue changing the Tank#?" + vbCrLf + _
                                '        "If you proceed, it will delete existing information", MsgBoxStyle.YesNo, "MUSTER CP Reading Warning")
                                '    If result = MsgBoxResult.No Then
                                '        row.Cells("Tank#").CancelUpdate()
                                '        Exit Select
                                '    End If
                                'End If
                                'cpReadingInfo.Reset()
                                'cpReadingInfo.TankPipeID = row.Cells("Tank#").ValueListResolved.GetValue(row.Cells("Tank#").ValueListResolved.SelectedItemIndex)
                                'cpReadingInfo.TankPipeIndex = row.Cells("Tank#").ValueListResolved.GetText(row.Cells("Tank#").ValueListResolved.SelectedItemIndex)
                                'row.Cells("TANK_INDEX").Value = cpReadingInfo.TankPipeIndex
                                '' archive so when you reset, the tank / pipe / term id/index is not lost
                                '' there is no other data entered before tank / pipe / term is selected
                                'cpReadingInfo.Archive()
                            Case "Pipe#"
                                If oInspection.CheckListMaster.InspectionCPReadings.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'cpReadingInfo = oInspection.InspectionInfo.CPReadingsCollection.Item(row.Cells("ID").Value.ToString)
                                If row.IsAddRow Then
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeID = row.Cells("Pipe#").ValueListResolved.GetValue(row.Cells("Pipe#").ValueListResolved.SelectedItemIndex)
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex = row.Cells("Pipe#").ValueListResolved.GetText(row.Cells("Pipe#").ValueListResolved.SelectedItemIndex)
                                    row.Cells("PIPE_INDEX").Value = oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex
                                    ' archive so when you reset, the tank / pipe / term id/index is not lost
                                    ' there is no other data entered before tank / pipe / term is selected
                                    oInspection.CheckListMaster.InspectionCPReadings.Archive()
                                Else
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeID = row.Cells("Pipe#").ValueListResolved.GetValue(row.Cells("Pipe#").ValueListResolved.SelectedItemIndex)
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex = row.Cells("Pipe#").ValueListResolved.GetText(row.Cells("Pipe#").ValueListResolved.SelectedItemIndex)
                                    row.Cells("PIPE_INDEX").Value = oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex
                                End If
                                ' need to update fuel type with respect to Tank#
                                row.Cells("Fuel Type").Value = oInspection.CheckListMaster.PipeFuelType.Item(oInspection.CheckListMaster.InspectionCPReadings.TankPipeID)
                                ' need to update the citaion with correct Pipe#
                                UpdateCitationCCATForCP(row, checkList)

                                'If cpReadingInfo.TankPipeID <> 0 And Not row.IsAddRow Then
                                '    Dim result As MsgBoxResult = MsgBox("Do you want to continue changing the Pipe#?" + vbCrLf + _
                                '        "If you proceed, it will delete existing information", MsgBoxStyle.YesNo, "MUSTER CP Reading Warning")
                                '    If result = MsgBoxResult.No Then
                                '        row.Cells("Pipe#").CancelUpdate()
                                '        Exit Select
                                '    End If
                                'End If
                                'cpReadingInfo.Reset()
                                'cpReadingInfo.TankPipeID = row.Cells("Pipe#").ValueListResolved.GetValue(row.Cells("Pipe#").ValueListResolved.SelectedItemIndex)
                                'cpReadingInfo.TankPipeIndex = row.Cells("Pipe#").ValueListResolved.GetText(row.Cells("Pipe#").ValueListResolved.SelectedItemIndex)
                                'row.Cells("PIPE_INDEX").Value = cpReadingInfo.TankPipeIndex
                                '' archive so when you reset, the tank / pipe / term id/index is not lost
                                '' there is no other data entered before tank / pipe / term is selected
                                'cpReadingInfo.Archive()
                            Case "Term#"
                                If oInspection.CheckListMaster.InspectionCPReadings.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'cpReadingInfo = oInspection.InspectionInfo.CPReadingsCollection.Item(row.Cells("ID").Value.ToString)
                                If row.IsAddRow Then
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeID = row.Cells("Term#").ValueListResolved.GetValue(row.Cells("Term#").ValueListResolved.SelectedItemIndex)
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex = row.Cells("Term#").ValueListResolved.GetText(row.Cells("Term#").ValueListResolved.SelectedItemIndex)
                                    row.Cells("TERM_INDEX").Value = oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex
                                    ' archive so when you reset, the tank / pipe / term id/index is not lost
                                    ' there is no other data entered before tank / pipe / term is selected
                                    oInspection.CheckListMaster.InspectionCPReadings.Archive()
                                Else
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeID = row.Cells("Term#").ValueListResolved.GetValue(row.Cells("Term#").ValueListResolved.SelectedItemIndex)
                                    oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex = row.Cells("Term#").ValueListResolved.GetText(row.Cells("Term#").ValueListResolved.SelectedItemIndex)
                                    row.Cells("TERM_INDEX").Value = oInspection.CheckListMaster.InspectionCPReadings.TankPipeIndex
                                End If
                                ' need to update fuel type with respect to Tank#
                                row.Cells("Fuel Type").Value = oInspection.CheckListMaster.PipeFuelType.Item(oInspection.CheckListMaster.InspectionCPReadings.TankPipeID)
                                ' need to update the citaion with correct Term#
                                UpdateCitationCCATForCP(row, checkList)

                                'If cpReadingInfo.TankPipeID <> 0 And Not row.IsAddRow Then
                                '    Dim result As MsgBoxResult = MsgBox("Do you want to continue changing the Term#?" + vbCrLf + _
                                '        "If you proceed, it will delete existing information", MsgBoxStyle.YesNo, "MUSTER CP Reading Warning")
                                '    If result = MsgBoxResult.No Then
                                '        row.Cells("Term#").CancelUpdate()
                                '        Exit Select
                                '    End If
                                'End If
                                'cpReadingInfo.Reset()
                                'cpReadingInfo.TankPipeID = row.Cells("Term#").ValueListResolved.GetValue(row.Cells("Term#").ValueListResolved.SelectedItemIndex)
                                'cpReadingInfo.TankPipeIndex = row.Cells("Term#").ValueListResolved.GetText(row.Cells("Term#").ValueListResolved.SelectedItemIndex)
                                'row.Cells("TERM_INDEX").Value = cpReadingInfo.TankPipeIndex
                                '' archive so when you reset, the tank / pipe / term id/index is not lost
                                '' there is no other data entered before tank / pipe / term is selected
                                'cpReadingInfo.Archive()
                        End Select
                    Case ugTankLeak.Name, ugPipeLeak.Name, ugMW.Name
                        Select Case ugCell
                            Case "Well#", "Well Depth", "Depth to" + vbCrLf + "Water", "Depth to" + vbCrLf + "Slots", "Inspector's Observations"
                                If oInspection.CheckListMaster.InspectionMonitorWells.ID <> row.Cells("ID").Value Then
                                    oInspection.CheckListMaster.InspectionMonitorWells.Retrieve(oInspection.InspectionInfo, row.Cells("ID").Value)
                                End If
                                'mwInfo = oInspection.InspectionInfo.MonitorWellsCollection.Item(row.Cells("ID").Value.ToString)
                                Select Case ugCell
                                    Case "Well#"
                                        If IsNumeric(row.Cells("Well#").EditorResolved.Value) Then
                                            oInspection.CheckListMaster.InspectionMonitorWells.WellNumber = row.Cells("Well#").EditorResolved.Value
                                        Else
                                            oInspection.CheckListMaster.InspectionMonitorWells.WellNumber = 0
                                        End If
                                        'Case "Well Depth"
                                        '    If IsNumeric(row.Cells("Well Depth").EditorResolved.Value) Then
                                        '        oInspection.CheckListMaster.InspectionMonitorWells.WellDepth = row.Cells("Well Depth").EditorResolved.Value
                                        '    Else
                                        '        oInspection.CheckListMaster.InspectionMonitorWells.WellDepth = 0.0
                                        '    End If
                                        'Case "Depth to" + vbCrLf + "Water"
                                        '    If IsNumeric(row.Cells("Depth to" + vbCrLf + "Water").EditorResolved.Value) Then
                                        '        oInspection.CheckListMaster.InspectionMonitorWells.DepthToWater = row.Cells("Depth to" + vbCrLf + "Water").EditorResolved.Value
                                        '    Else
                                        '        oInspection.CheckListMaster.InspectionMonitorWells.DepthToWater = 0.0
                                        '    End If
                                        'Case "Depth to" + vbCrLf + "Slots"
                                        '    If IsNumeric(row.Cells("Depth to" + vbCrLf + "Slots").EditorResolved.Value) Then
                                        '        oInspection.CheckListMaster.InspectionMonitorWells.DepthToSlots = row.Cells("Depth to" + vbCrLf + "Slots").EditorResolved.Value
                                        '    Else
                                        '        oInspection.CheckListMaster.InspectionMonitorWells.DepthToSlots = 0.0
                                        '    End If
                                End Select
                                row.Cells("Well#").Value = oInspection.CheckListMaster.InspectionMonitorWells.WellNumber
                                'row.Cells("Well Depth").Value = oInspection.CheckListMaster.InspectionMonitorWells.WellDepth
                                'row.Cells("Depth to" + vbCrLf + "Water").Value = oInspection.CheckListMaster.InspectionMonitorWells.DepthToWater
                                'row.Cells("Depth to" + vbCrLf + "Slots").Value = oInspection.CheckListMaster.InspectionMonitorWells.DepthToSlots
                                oInspection.CheckListMaster.InspectionMonitorWells.WellDepth = row.Cells("Well Depth").Text
                                oInspection.CheckListMaster.InspectionMonitorWells.DepthToWater = row.Cells("Depth to" + vbCrLf + "Water").Text
                                oInspection.CheckListMaster.InspectionMonitorWells.DepthToSlots = row.Cells("Depth to" + vbCrLf + "Slots").Text
                                oInspection.CheckListMaster.InspectionMonitorWells.InspectorsObservations = row.Cells("Inspector's Observations").Text
                                ' if well# is updated, need to update the citaion with correct Well#
                                If ugCell = "Well#" Then
                                    UpdateCitationCCATForMWell(row, checkList)
                                End If
                                SetupTankPipeLeakRow(row)
                        End Select
                End Select
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            btnSave.Enabled = oInspection.IsDirty Or oInspection.CheckListMaster.colIsDirty
        End Try
    End Sub
    Private Sub UpdateCitationCCATForCP(ByVal row As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal checklist As MUSTER.Info.InspectionChecklistMasterInfo)
        Try
            Dim strUpdatedCCAT As String = ""
            Dim slCCAT As New SortedList
            Dim deleteCitation As Boolean = True
            Dim deleteDiscrep As Boolean = True

            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In row.ParentCollection
                ' current row is not the same row being updated
                'If ugRow.Cells("ID").Value <> row.Cells("ID").Value Then
                ' if any other row belonging to the updating row's parent
                ' has "Fail" or "Incon" set to true, set delete variables to false
                If ugRow.Cells("Fail").Value Or ugRow.Cells("Incon").Value Then
                    If ugRow.Cells("Line#").Text.StartsWith("3.5.4.") Then
                        'strUpdatedCCAT += "T" + ugRow.Cells("Tank#").Text + ", "
                        If Not slCCAT.Contains(ugRow.Cells("Tank#").Value) Then
                            slCCAT.Add(ugRow.Cells("Tank#").Value, "T" + ugRow.Cells("Tank#").Text)
                        End If
                    ElseIf ugRow.Cells("Line#").Text.StartsWith("3.6.3.") Then
                        'strUpdatedCCAT += "P" + ugRow.Cells("Pipe#").Text + ", "
                        If Not slCCAT.Contains(ugRow.Cells("Pipe#").Value) Then
                            slCCAT.Add(ugRow.Cells("Pipe#").Value, "P" + ugRow.Cells("Pipe#").Text)
                        End If
                    ElseIf ugRow.Cells("Line#").Text.StartsWith("3.7.6.") Then
                        'strUpdatedCCAT += "TP" + ugRow.Cells("Term#").Text + ", "
                        If Not slCCAT.Contains(ugRow.Cells("Term#").Value) Then
                            slCCAT.Add(ugRow.Cells("Term#").Value, "TP" + ugRow.Cells("Term#").Text)
                        End If
                    Else
                        deleteCitation = False
                    End If
                    deleteDiscrep = False
                End If
                'End If
            Next

            For i As Integer = 0 To slCCAT.Count - 1
                strUpdatedCCAT += slCCAT.GetByIndex(i).ToString + ", "
            Next
            If strUpdatedCCAT <> String.Empty Then
                strUpdatedCCAT = strUpdatedCCAT.Trim.TrimEnd(",")
            End If

            If deleteCitation Then
                For Each citation As MUSTER.Info.InspectionCitationInfo In oInspection.InspectionInfo.CitationsCollection.Values
                    If citation.QuestionID = checklist.ID Then
                        If oInspection.CheckListMaster.InspectionCitation.ID <> citation.ID Then
                            oInspection.CheckListMaster.InspectionCitation.Retrieve(oInspection.InspectionInfo, citation.ID)
                        End If
                        If strUpdatedCCAT <> String.Empty Then
                            'citation.Reset()
                            oInspection.CheckListMaster.InspectionCitation.CCAT = strUpdatedCCAT
                            oInspection.CheckListMaster.InspectionCitation.Deleted = False
                        Else
                            oInspection.CheckListMaster.InspectionCitation.Reset()
                            oInspection.CheckListMaster.InspectionCitation.CCAT = String.Empty
                            oInspection.CheckListMaster.InspectionCitation.Deleted = True
                        End If
                        Exit For
                    End If
                Next
            End If

            If deleteDiscrep Then
                For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.InspectionInfo.DiscrepsCollection.Values
                    If discrep.QuestionID = checklist.ID Then
                        If oInspection.CheckListMaster.InspectionDiscrep.ID <> discrep.ID Then
                            oInspection.CheckListMaster.InspectionDiscrep.Retrieve(oInspection.InspectionInfo, discrep.ID)
                        End If
                        oInspection.CheckListMaster.InspectionDiscrep.Reset()
                        oInspection.CheckListMaster.InspectionDiscrep.Deleted = True
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub UpdateCitationCCATForMWell(ByVal row As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal checklist As MUSTER.Info.InspectionChecklistMasterInfo)
        Try
            Dim strUpdatedCCAT As String = ""
            Dim slCCAT As New SortedList
            Dim deleteCitation As Boolean = True
            Dim deleteDiscrep As Boolean = True

            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In row.ParentCollection
                ' current row is not the same row being updated
                'If ugRow.Cells("ID").Value <> row.Cells("ID").Value Then
                ' if any other row belonging to the updating row's parent
                ' has "No" set to true, set delete variables to false
                If ugRow.Cells("Surface Sealed" + vbCrLf + "No").Value Or _
                    ugRow.Cells("Well Caps" + vbCrLf + "No").Value Then

                    If Not slCCAT.Contains(ugRow.Cells("Well#").Value) Then
                        slCCAT.Add(ugRow.Cells("Well#").Value, ugRow.Cells("Well#").Value.ToString)
                    End If

                    'deleteCitation = False
                    deleteDiscrep = False
                End If
                'End If
            Next

            For i As Integer = 0 To slCCAT.Count - 1
                strUpdatedCCAT += slCCAT.GetByIndex(i).ToString + ", "
            Next
            If strUpdatedCCAT <> String.Empty Then
                strUpdatedCCAT = strUpdatedCCAT.Trim.TrimEnd(",")
            End If

            If deleteCitation Then
                For Each citation As MUSTER.Info.InspectionCitationInfo In oInspection.InspectionInfo.CitationsCollection.Values
                    If citation.QuestionID = checklist.ID Then
                        If oInspection.CheckListMaster.InspectionCitation.ID <> citation.ID Then
                            oInspection.CheckListMaster.InspectionCitation.Retrieve(oInspection.InspectionInfo, citation.ID)
                        End If
                        If strUpdatedCCAT <> String.Empty Then
                            'citation.Reset()
                            oInspection.CheckListMaster.InspectionCitation.CCAT = strUpdatedCCAT
                            oInspection.CheckListMaster.InspectionCitation.Deleted = False
                        Else
                            oInspection.CheckListMaster.InspectionCitation.Reset()
                            oInspection.CheckListMaster.InspectionCitation.CCAT = String.Empty
                            oInspection.CheckListMaster.InspectionCitation.Deleted = True
                        End If
                        Exit For
                    End If
                Next
            End If

            If deleteDiscrep Then
                For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.InspectionInfo.DiscrepsCollection.Values
                    If discrep.QuestionID = checklist.ID Then
                        If oInspection.CheckListMaster.InspectionDiscrep.ID <> discrep.ID Then
                            oInspection.CheckListMaster.InspectionDiscrep.Retrieve(oInspection.InspectionInfo, discrep.ID)
                        End If
                        oInspection.CheckListMaster.InspectionDiscrep.Reset()
                        oInspection.CheckListMaster.InspectionDiscrep.Deleted = True
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ugReg_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugReg.BeforeCellActivate
        BeforeCellActivate(ugReg, e)
    End Sub
    Private Sub ugSpill_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugSpill.BeforeCellActivate
        BeforeCellActivate(ugSpill, e)
    End Sub
    Private Sub ugCP_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugCP.BeforeCellActivate
        BeforeCellActivate(ugCP, e)
    End Sub
    Private Sub ugTankLeak_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugTankLeak.BeforeCellActivate
        BeforeCellActivate(ugTankLeak, e)
    End Sub
    Private Sub ugPipeLeak_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugPipeLeak.BeforeCellActivate
        BeforeCellActivate(ugPipeLeak, e)
    End Sub
    Private Sub ugCatLeak_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugCatLeak.BeforeCellActivate
        BeforeCellActivate(ugCatLeak, e)
    End Sub
    Private Sub ugVisual_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugVisual.BeforeCellActivate
        BeforeCellActivate(ugVisual, e)
    End Sub
    Private Sub ugTOS_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugTOS.BeforeCellActivate
        BeforeCellActivate(ugTOS, e)
    End Sub
    Private Sub ugSOC_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugSOC.BeforeCellActivate
        BeforeCellActivate(ugSOC, e)
    End Sub
    Private Sub ugMW_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugMW.BeforeCellActivate
        BeforeCellActivate(ugMW, e)
    End Sub
    Private Sub BeforeCellActivate(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs)
        Dim bolSetCellColor As Boolean = False
        Try
            If e.Cell.Row.Band.Index = 0 Then
                If "Yes".Equals(e.Cell.Column.Key) Then
                    bolSetCellColor = True
                ElseIf "No".Equals(e.Cell.Column.Key) Then
                    bolSetCellColor = True
                End If
            Else
                Select Case ug.Name
                    Case ugCP.Name
                        If e.Cell.Row.Band.Index = 2 Or e.Cell.Row.Band.Index = 6 Or e.Cell.Row.Band.Index = 10 Then
                            If "Yes".Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            ElseIf "No".Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            End If
                        ElseIf e.Cell.Row.Band.Index = 3 Or e.Cell.Row.Band.Index = 7 Or e.Cell.Row.Band.Index = 11 Then
                            If "Galvanic".Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            ElseIf "Impressed Current".Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            End If
                        ElseIf e.Cell.Row.Band.Index = 5 Or e.Cell.Row.Band.Index = 9 Or e.Cell.Row.Band.Index = 13 Then
                            If "Pass".Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            ElseIf "Fail".Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            ElseIf "Incon".Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            End If
                        End If
                    Case ugTankLeak.Name, ugPipeLeak.Name, ugMW.Name
                        If e.Cell.Row.Band.Index = 1 Then
                            If ("Surface Sealed" + vbCrLf + "Yes").Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            ElseIf ("Surface Sealed" + vbCrLf + "No").Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            ElseIf ("Well Caps" + vbCrLf + "Yes").Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            ElseIf ("Well Caps" + vbCrLf + "No").Equals(e.Cell.Column.Key) Then
                                bolSetCellColor = True
                            End If
                        End If
                End Select
            End If

            If bolSetCellColor Then
                e.Cell.Appearance.BackColor = Color.Blue
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugReg_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugReg.BeforeCellDeactivate
        BeforeCellDeactivate(ugReg)
    End Sub
    Private Sub ugSpill_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugSpill.BeforeCellDeactivate
        BeforeCellDeactivate(ugSpill)
    End Sub
    Private Sub ugCP_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugCP.BeforeCellDeactivate
        BeforeCellDeactivate(ugCP)
    End Sub
    Private Sub ugTankLeak_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugTankLeak.BeforeCellDeactivate
        BeforeCellDeactivate(ugTankLeak)
    End Sub
    Private Sub ugPipeLeak_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugPipeLeak.BeforeCellDeactivate
        BeforeCellDeactivate(ugPipeLeak)
    End Sub
    Private Sub ugCatLeak_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugCatLeak.BeforeCellDeactivate
        BeforeCellDeactivate(ugCatLeak)
    End Sub
    Private Sub ugVisual_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugVisual.BeforeCellDeactivate
        BeforeCellDeactivate(ugVisual)
    End Sub
    Private Sub ugTOS_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugTOS.BeforeCellDeactivate
        BeforeCellDeactivate(ugTOS)
    End Sub
    Private Sub ugSOC_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugSOC.BeforeCellDeactivate
        BeforeCellDeactivate(ugSOC)
    End Sub
    Private Sub ugMW_BeforeCellDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugMW.BeforeCellDeactivate
        BeforeCellDeactivate(ugMW)
    End Sub
    Private Sub BeforeCellDeactivate(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim bolReSetCellColor As Boolean = False
        Try
            If Not ug.ActiveCell Is Nothing Then
                If ug.ActiveCell.Row.Band.Index = 0 Then
                    If "Yes".Equals(ug.ActiveCell.Column.Key) Then
                        bolReSetCellColor = True
                    ElseIf "No".Equals(ug.ActiveCell.Column.Key) Then
                        bolReSetCellColor = True
                    End If
                Else
                    Select Case ug.Name
                        Case ugCP.Name
                            If ug.ActiveCell.Row.Band.Index = 2 Or ug.ActiveCell.Row.Band.Index = 6 Or ug.ActiveCell.Row.Band.Index = 10 Then
                                If "Yes".Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                ElseIf "No".Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                End If
                            ElseIf ug.ActiveCell.Row.Band.Index = 3 Or ug.ActiveCell.Row.Band.Index = 7 Or ug.ActiveCell.Row.Band.Index = 11 Then
                                If "Galvanic".Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                ElseIf "Impressed Current".Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                End If
                            ElseIf ug.ActiveCell.Row.Band.Index = 5 Or ug.ActiveCell.Row.Band.Index = 9 Or ug.ActiveCell.Row.Band.Index = 13 Then
                                If "Pass".Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                ElseIf "Fail".Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                ElseIf "Incon".Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                End If
                            End If
                        Case ugTankLeak.Name, ugPipeLeak.Name, ugMW.Name
                            If ug.ActiveCell.Row.Band.Index = 1 Then
                                If ("Surface Sealed" + vbCrLf + "Yes").Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                ElseIf ("Surface Sealed" + vbCrLf + "No").Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                ElseIf ("Well Caps" + vbCrLf + "Yes").Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                ElseIf ("Well Caps" + vbCrLf + "No").Equals(ug.ActiveCell.Column.Key) Then
                                    bolReSetCellColor = True
                                End If
                            End If
                    End Select
                End If

                If bolReSetCellColor Then
                    ug.ActiveCell.Appearance.BackColor = Color.White
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugReg_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugReg.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugReg.ActiveCell Is Nothing Then
                If "CCAT".Equals(ugReg.ActiveCell.Column.Key) Then
                    If ugReg.ActiveRow.Cells("No").Value = True And ugReg.ActiveRow.Cells("CCAT").Value <> String.Empty Then
                        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
                        'Dim ccat As MUSTER.Info.InspectionCCATInfo
                        checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(ugReg.ActiveRow.Cells("QUESTION_ID").Value)
                        ' CCAT
                        If checkList.CCAT Then
                            ' reset ccat
                            'For Each ccat In oInspection.InspectionInfo.CCATsCollection.Values
                            '    If ccat.QuestionID = checkList.ID Then
                            '        ccat.Reset()
                            '    End If
                            'Next
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugReg.ActiveRow, bolReadOnly, , , True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugSpill_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugSpill.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugSpill.ActiveCell Is Nothing Then
                If "CCAT".Equals(ugSpill.ActiveCell.Column.Key) Then
                    If ugSpill.ActiveRow.Cells("No").Value = True And ugSpill.ActiveRow.Cells("CCAT").Value <> String.Empty Then
                        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
                        'Dim ccat As MUSTER.Info.InspectionCCATInfo
                        checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(ugSpill.ActiveRow.Cells("QUESTION_ID").Value)
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugSpill.ActiveRow, bolReadOnly, , , True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCP_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugCP.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugCP.ActiveCell Is Nothing Then
                Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
                checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(ugCP.ActiveRow.Cells("QUESTION_ID").Value)
                If "CCAT".Equals(ugCP.ActiveCell.Column.Key) Then
                    If ugCP.ActiveRow.Cells("No").Value = True And ugCP.ActiveRow.Cells("CCAT").Value <> String.Empty Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugCP.ActiveRow, bolReadOnly, , , True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                ElseIf "Fail".Equals(ugCP.ActiveCell.Column.Key) Then
                    If ugCP.ActiveRow.Cells("PASSFAILINCON").Value <> 1 And ugCP.ActiveRow.Cells("Fail").Value = True Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugCP.ActiveRow, bolReadOnly, ugCP.Name, ugCP.ActiveCell.Column.Key, True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                ElseIf "Incon".Equals(ugCP.ActiveCell.Column.Key) Then
                    If ugCP.ActiveRow.Cells("PASSFAILINCON").Value <> 1 And ugCP.ActiveRow.Cells("Incon").Value = True Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugCP.ActiveRow, bolReadOnly, ugCP.Name, ugCP.ActiveCell.Column.Key, True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTankLeak_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugTankLeak.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugTankLeak.ActiveCell Is Nothing Then
                Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
                checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(ugTankLeak.ActiveRow.Cells("QUESTION_ID").Value)
                If "CCAT".Equals(ugTankLeak.ActiveCell.Column.Key) Then
                    If ugTankLeak.ActiveRow.Cells("No").Value = True And ugTankLeak.ActiveRow.Cells("CCAT").Value <> String.Empty Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugTankLeak.ActiveRow, bolReadOnly, , , True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                ElseIf ("Surface Sealed" + vbCrLf + "No").Equals(ugTankLeak.ActiveCell.Column.Key) Then
                    If ugTankLeak.ActiveRow.Cells("SURFACE_SEALED").Value <> 1 And ugTankLeak.ActiveRow.Cells("Surface Sealed" + vbCrLf + "No").Value = True Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugTankLeak.ActiveRow, bolReadOnly, ugTankLeak.Name, ugTankLeak.ActiveCell.Column.Key, True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                ElseIf ("Well Caps" + vbCrLf + "No").Equals(ugTankLeak.ActiveCell.Column.Key) Then
                    If ugTankLeak.ActiveRow.Cells("WELL_CAPS").Value <> 1 And ugTankLeak.ActiveRow.Cells("Well Caps" + vbCrLf + "No").Value = True Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugTankLeak.ActiveRow, bolReadOnly, ugTankLeak.Name, ugTankLeak.ActiveCell.Column.Key, True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugPipeLeak_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugPipeLeak.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugPipeLeak.ActiveCell Is Nothing Then
                Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
                checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(ugPipeLeak.ActiveRow.Cells("QUESTION_ID").Value)
                If "CCAT".Equals(ugPipeLeak.ActiveCell.Column.Key) Then
                    If ugPipeLeak.ActiveRow.Cells("No").Value = True And ugPipeLeak.ActiveRow.Cells("CCAT").Value <> String.Empty Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugPipeLeak.ActiveRow, bolReadOnly, , , True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                ElseIf ("Surface Sealed" + vbCrLf + "No").Equals(ugPipeLeak.ActiveCell.Column.Key) Then
                    If ugPipeLeak.ActiveRow.Cells("SURFACE_SEALED").Value <> 1 And ugPipeLeak.ActiveRow.Cells("Surface Sealed" + vbCrLf + "No").Value = True Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugPipeLeak.ActiveRow, bolReadOnly, ugPipeLeak.Name, ugPipeLeak.ActiveCell.Column.Key, True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                ElseIf ("Well Caps" + vbCrLf + "No").Equals(ugPipeLeak.ActiveCell.Column.Key) Then
                    If ugPipeLeak.ActiveRow.Cells("WELL_CAPS").Value <> 1 And ugPipeLeak.ActiveRow.Cells("Well Caps" + vbCrLf + "No").Value = True Then
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugPipeLeak.ActiveRow, bolReadOnly, ugPipeLeak.Name, ugPipeLeak.ActiveCell.Column.Key, True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCatLeak_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugCatLeak.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugCatLeak.ActiveCell Is Nothing Then
                If "CCAT".Equals(ugCatLeak.ActiveCell.Column.Key) Then
                    If ugCatLeak.ActiveRow.Cells("No").Value = True And ugCatLeak.ActiveRow.Cells("CCAT").Value <> String.Empty Then
                        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
                        checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(ugCatLeak.ActiveRow.Cells("QUESTION_ID").Value)
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugCatLeak.ActiveRow, bolReadOnly, , , True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugVisual_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugVisual.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugVisual.ActiveCell Is Nothing Then
                If "CCAT".Equals(ugVisual.ActiveCell.Column.Key) Then
                    If ugVisual.ActiveRow.Cells("No").Value = True And ugVisual.ActiveRow.Cells("CCAT").Value <> String.Empty Then
                        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
                        checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(ugVisual.ActiveRow.Cells("QUESTION_ID").Value)
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugVisual.ActiveRow, bolReadOnly, , , True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTOS_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugTOS.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugTOS.ActiveCell Is Nothing Then
                If "CCAT".Equals(ugTOS.ActiveCell.Column.Key) Then
                    If ugTOS.ActiveRow.Cells("No").Value = True And ugTOS.ActiveRow.Cells("CCAT").Value <> String.Empty Then
                        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
                        checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(ugTOS.ActiveRow.Cells("QUESTION_ID").Value)
                        ' CCAT
                        If checkList.CCAT Then
                            Dim frmCCAT As CCAT
                            frmCCAT = New CCAT(oInspection, ugTOS.ActiveRow, bolReadOnly, , , True)
                            frmCCAT.ShowDialog()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub TankPipeLeakBeforeRowInsert(ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs, ByVal isTankBed As Boolean)
    '    Dim mw As MUSTER.Info.InspectionMonitorWellsInfo
    '    Try
    '        mw = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
    '            oInspection.ID, _
    '            CType(e.ParentRow.Cells("QUESTION_ID").Value, Int64), _
    '            False, _
    '            0, _
    '            0.0, _
    '            0.0, _
    '            0.0, _
    '            -1, _
    '            -1, _
    '            String.Empty, _
    '            False, _
    '            String.Empty, _
    '            dtNullDate, _
    '            String.Empty, _
    '            dtNullDate, _
    '            IIf(isTankBed, oInspection.CheckListMaster.MaxTankBedMWLineNum, oInspection.CheckListMaster.MaxLineMWLineNum))
    '        oInspection.CheckListMaster.InspectionMonitorWells.Add(mw)
    '        If isTankBed Then
    '            oInspection.CheckListMaster.MaxTankBedMWLineNum += 1
    '        Else
    '            oInspection.CheckListMaster.MaxLineMWLineNum += 1
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        btnSave.Enabled = oInspection.IsDirty Or oInspection.CheckListMaster.colIsDirty
    '    End Try
    'End Sub
    'Private Sub TankPipeLeakAfterRowInsert(ByRef ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow)
    '    Dim mw As MUSTER.BusinessLogic.pInspectionMonitorWells
    '    Try
    '        mw = oInspection.CheckListMaster.InspectionMonitorWells
    '        ugRow.Cells("Line#").Value = ugRow.ParentRow.Cells("Line#").Value + "." + mw.LineNumber.ToString
    '        ugRow.Cells("Well#").Value = mw.WellNumber
    '        ugRow.Cells("Well Depth").Value = mw.WellDepth
    '        ugRow.Cells("Depth to" + vbCrLf + "Water").Value = mw.DepthToWater
    '        ugRow.Cells("Depth to" + vbCrLf + "Slots").Value = mw.DepthToSlots
    '        ugRow.Cells("Surface Sealed" + vbCrLf + "Yes").Value = IIf(mw.SurfaceSealed = 1, True, False)
    '        ugRow.Cells("Surface Sealed" + vbCrLf + "No").Value = IIf(mw.SurfaceSealed = 0, True, False)
    '        ugRow.Cells("Well Caps" + vbCrLf + "Yes").Value = IIf(mw.WellCaps = 1, True, False)
    '        ugRow.Cells("Well Caps" + vbCrLf + "No").Value = IIf(mw.WellCaps = 0, True, False)
    '        ugRow.Cells("ID").Value = mw.ID
    '        ugRow.Cells("INSPECTION_ID").Value = mw.InspectionID
    '        ugRow.Cells("QUESTION_ID").Value = mw.QuestionID
    '        ugRow.Cells("TANK_LINE").Value = mw.TankLine
    '        ugRow.Cells("SURFACE_SEALED").Value = mw.SurfaceSealed
    '        ugRow.Cells("WELL_CAPS").Value = mw.WellCaps
    '        ugRow.Cells("CITATION").Value = ugRow.ParentRow.Cells("CITATION").Value
    '        ugRow.Cells("LINE_NUMBER").Value = mw.LineNumber
    '        SetupTankPipeLeakRow(ugRow)
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub ugTankLeak_BeforeRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs) Handles ugTankLeak.BeforeRowInsert
    '    Try
    '        TankPipeLeakBeforeRowInsert(e, True)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugTankLeak_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugTankLeak.AfterRowInsert
    '    Try
    '        TankPipeLeakAfterRowInsert(e.Row)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugPipeLeak_BeforeRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs) Handles ugPipeLeak.BeforeRowInsert
    '    Try
    '        TankPipeLeakBeforeRowInsert(e, False)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugPipeLeak_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugPipeLeak.AfterRowInsert
    '    Try
    '        TankPipeLeakAfterRowInsert(e.Row)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub btnAddTankMW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddTankMW.Click
        Try
            Cursor.Current = Cursors.AppStarting
            AddTankPipeLeakMW(ugTankLeak)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub btnAddPipeMW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPipeMW.Click
        Try
            Cursor.Current = Cursors.AppStarting
            AddTankPipeLeakMW(ugPipeLeak)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub btnMWAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMWAdd.Click
        Try
            Cursor.Current = Cursors.AppStarting
            AddTankPipeLeakMW(ugMW)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub AddTankPipeLeakMW(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        Dim mw As MUSTER.Info.InspectionMonitorWellsInfo
        Dim parentLineNum, childLineNum As String
        Dim bolFindRow As Boolean = False
        Try
            Select Case ug.Name
                Case ugTankLeak.Name
                    parentLineNum = "4.2.8"
                    childLineNum = "4.2.8." + oInspection.CheckListMaster.MaxTankBedMWLineNum.ToString
                Case ugPipeLeak.Name
                    parentLineNum = "5.2.8"
                    childLineNum = "5.2.8." + oInspection.CheckListMaster.MaxLineMWLineNum.ToString
                Case ugMW.Name
                    parentLineNum = "11"
                    childLineNum = "11." + oInspection.CheckListMaster.MaxTankLineMWLineNum.ToString
            End Select
            If ug.ActiveRow Is Nothing Then
                bolFindRow = True
            ElseIf ug.ActiveRow.Cells("Line#").Value <> parentLineNum Then
                bolFindRow = True
            End If
            If bolFindRow Then
                For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ug.Rows
                    If Not dr.Cells("Line#").Value Is Nothing Then
                        If dr.Cells("Line#").Value = parentLineNum Then
                            ug.ActiveRow = dr
                            Exit For
                        End If
                    End If
                Next
            End If
            ug.DisplayLayout.Bands(1).AddNew()
            ' after addnew(), active row is the new row added
            Select Case parentLineNum 'ugCP.ActiveRow.ParentRow.Cells("Line#").Value
                Case "4.2.8"
                    mw = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                        oInspection.ID, _
                        CType(ug.ActiveRow.ParentRow.Cells("QUESTION_ID").Value, Int64), _
                        True, _
                        0, _
                        String.Empty, _
                        String.Empty, _
                        String.Empty, _
                        -1, _
                        -1, _
                        String.Empty, _
                        False, _
                        String.Empty, _
                        dtNullDate, _
                        String.Empty, _
                        dtNullDate, _
                        oInspection.CheckListMaster.MaxTankBedMWLineNum)
                    oInspection.CheckListMaster.InspectionMonitorWells.Add(mw)
                    oInspection.CheckListMaster.MaxTankBedMWLineNum += 1
                Case "5.2.8"
                    mw = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                        oInspection.ID, _
                        CType(ug.ActiveRow.ParentRow.Cells("QUESTION_ID").Value, Int64), _
                        False, _
                        0, _
                        String.Empty, _
                        String.Empty, _
                        String.Empty, _
                        -1, _
                        -1, _
                        String.Empty, _
                        False, _
                        String.Empty, _
                        dtNullDate, _
                        String.Empty, _
                        dtNullDate, _
                        oInspection.CheckListMaster.MaxLineMWLineNum)
                    oInspection.CheckListMaster.InspectionMonitorWells.Add(mw)
                    oInspection.CheckListMaster.MaxLineMWLineNum += 1
                Case "11"
                    mw = New MUSTER.Info.InspectionMonitorWellsInfo(0, _
                        oInspection.ID, _
                        CType(ug.ActiveRow.ParentRow.Cells("QUESTION_ID").Value, Int64), _
                        False, _
                        0, _
                        String.Empty, _
                        String.Empty, _
                        String.Empty, _
                        -1, _
                        -1, _
                        String.Empty, _
                        False, _
                        String.Empty, _
                        dtNullDate, _
                        String.Empty, _
                        dtNullDate, _
                        oInspection.CheckListMaster.MaxLineMWLineNum)
                    oInspection.CheckListMaster.InspectionMonitorWells.Add(mw)
                    oInspection.CheckListMaster.MaxTankLineMWLineNum += 1
            End Select
            ug.ActiveRow.Cells("Line#").Value = childLineNum
            ug.ActiveRow.Cells("Well#").Value = mw.WellNumber
            'ug.ActiveRow.Cells("Well Depth").Value = mw.WellDepth
            'ug.ActiveRow.Cells("Depth to" + vbCrLf + "Water").Value = mw.DepthToWater
            'ug.ActiveRow.Cells("Depth to" + vbCrLf + "Slots").Value = mw.DepthToSlots
            ug.ActiveRow.Cells("Surface Sealed" + vbCrLf + "Yes").Value = IIf(mw.SurfaceSealed = 1, True, False)
            ug.ActiveRow.Cells("Surface Sealed" + vbCrLf + "No").Value = IIf(mw.SurfaceSealed = 0, True, False)
            ug.ActiveRow.Cells("Well Caps" + vbCrLf + "Yes").Value = IIf(mw.WellCaps = 1, True, False)
            ug.ActiveRow.Cells("Well Caps" + vbCrLf + "No").Value = IIf(mw.WellCaps = 0, True, False)
            ug.ActiveRow.Cells("ID").Value = mw.ID
            ug.ActiveRow.Cells("INSPECTION_ID").Value = mw.InspectionID
            ug.ActiveRow.Cells("QUESTION_ID").Value = mw.QuestionID
            ug.ActiveRow.Cells("TANK_LINE").Value = mw.TankLine
            ug.ActiveRow.Cells("SURFACE_SEALED").Value = mw.SurfaceSealed
            ug.ActiveRow.Cells("WELL_CAPS").Value = mw.WellCaps
            ug.ActiveRow.Cells("CITATION").Value = ug.ActiveRow.ParentRow.Cells("CITATION").Value
            ug.ActiveRow.Cells("LINE_NUMBER").Value = mw.LineNumber
            SetupTankPipeLeakRow(ug.ActiveRow)
            btnSave.Enabled = oInspection.IsDirty Or oInspection.CheckListMaster.colIsDirty

            ug.Focus()
            ug.ActiveRow.Cells("Line#").Activate()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugTankLeak_BeforeRowsDeleted(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles ugTankLeak.BeforeRowsDeleted
        ugTankPipeLeakBeforeRowsDeleted(sender, e)
    End Sub
    Private Sub ugPipeLeak_BeforeRowsDeleted(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles ugPipeLeak.BeforeRowsDeleted
        ugTankPipeLeakBeforeRowsDeleted(sender, e)
    End Sub
    Private Sub ugMW_BeforeRowsDeleted(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles ugMW.BeforeRowsDeleted
        ugTankPipeLeakBeforeRowsDeleted(sender, e)
    End Sub
    Private Sub ugTankPipeLeakBeforeRowsDeleted(ByVal sender As System.Object, ByRef e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs)
        'Dim mwells As MUSTER.Info.InspectionMonitorWellsInfo
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim msgBoxResult As Windows.Forms.DialogResult
        Try
            e.DisplayPromptMsg = False
            msgBoxResult = MessageBox.Show("You have selected " + _
                            e.Rows.Length.ToString + " " + _
                            IIf(e.Rows.Length > 1, "rows", "row") + " for deletion." + vbCrLf + _
                            "Choose Yes to delete the " + _
                            IIf(e.Rows.Length > 1, "rows", "row") + _
                            " or No to exit.", "Delete Rows", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If msgBoxResult = DialogResult.Yes Then
                Dim alRowIDs As New ArrayList
                Dim slRowIDs As New SortedList
                For Each dr In e.Rows
                    If Not slRowIDs.Contains(dr.Cells("LINE_NUMBER").Value) Then
                        slRowIDs.Add(dr.Cells("LINE_NUMBER").Value, dr)
                    End If
                    If Not alRowIDs.Contains(dr.Cells("ID").Value.ToString) Then
                        alRowIDs.Add(dr.Cells("ID").Value.ToString)
                    End If
                Next
                For i As Integer = slRowIDs.Count - 1 To 0 Step -1 'Each dr In e.Rows
                    'If Not alRowIDs.Contains(dr.Cells("ID").Value.ToString) Then
                    '    alRowIDs.Add(dr.Cells("ID").Value.ToString)
                    'End If
                    dr = slRowIDs.GetByIndex(i)
                    If oInspection.CheckListMaster.InspectionMonitorWells.ID <> dr.Cells("ID").Value Then
                        oInspection.CheckListMaster.InspectionMonitorWells.Retrieve(oInspection.InspectionInfo, dr.Cells("ID").Value)
                    End If
                    'mwells = oInspection.InspectionInfo.MonitorWellsCollection.Item(dr.Cells("ID").Value.ToString)
                    oInspection.CheckListMaster.InspectionMonitorWells.Deleted = True
                    If oInspection.CheckListMaster.InspectionMonitorWells.ID < 0 Then
                        oInspection.InspectionInfo.MonitorWellsCollection.Remove(oInspection.CheckListMaster.InspectionMonitorWells.ID)
                    End If
                    If dr.IsAddRow Then
                        If dr.Cells("Line#").Text.StartsWith("4.2.8") And (oInspection.CheckListMaster.MaxTankBedMWLineNum - 1) > 0 Then
                            oInspection.CheckListMaster.MaxTankBedMWLineNum -= 1
                        ElseIf dr.Cells("Line#").Text.StartsWith("5.2.8") And (oInspection.CheckListMaster.MaxLineMWLineNum - 1) > 0 Then
                            oInspection.CheckListMaster.MaxLineMWLineNum -= 1
                        ElseIf dr.Cells("Line#").Text.StartsWith("11") And (oInspection.CheckListMaster.MaxTankLineMWLineNum - 1) > 0 Then
                            oInspection.CheckListMaster.MaxTankLineMWLineNum -= 1
                        End If
                    Else
                        If dr.Cells("Line#").Text.StartsWith("4.2.8") Then
                            If dr.Cells("LINE_NUMBER").Value = oInspection.CheckListMaster.MaxTankBedMWLineNum - 1 And (oInspection.CheckListMaster.MaxTankBedMWLineNum - 1) > 0 Then
                                oInspection.CheckListMaster.MaxTankBedMWLineNum -= 1
                            End If
                        ElseIf dr.Cells("Line#").Text.StartsWith("5.2.8") Then
                            If dr.Cells("LINE_NUMBER").Value = oInspection.CheckListMaster.MaxLineMWLineNum - 1 And (oInspection.CheckListMaster.MaxLineMWLineNum - 1) > 0 Then
                                oInspection.CheckListMaster.MaxLineMWLineNum -= 1
                            End If
                        ElseIf dr.Cells("Line#").Text.StartsWith("11") Then
                            If dr.Cells("LINE_NUMBER").Value = oInspection.CheckListMaster.MaxTankLineMWLineNum - 1 And (oInspection.CheckListMaster.MaxTankLineMWLineNum - 1) > 0 Then
                                oInspection.CheckListMaster.MaxTankLineMWLineNum -= 1
                            End If
                        End If
                    End If
                Next
                Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo = oInspection.InspectionInfo.ChecklistMasterCollection.Item(e.Rows(0).Cells("QUESTION_ID").Value)
                If Not checkList Is Nothing Then
                    ' if there are any citations / discreps - delete
                    ' citation exists only if citation is not equal to -1
                    ' discrep exists only if DiscrepText is not empty
                    Dim strUpdatedCCAT As String = ""
                    Dim deleteCitation As Boolean = True
                    Dim deleteDiscrep As Boolean = True
                    If checkList.Citation <> -1 Or checkList.DiscrepText <> String.Empty Then
                        Dim slCCAT As New SortedList
                        deleteCitation = True
                        deleteDiscrep = True
                        For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows(0).ParentCollection
                            ' current row is one of the deleted rows
                            If Not alRowIDs.Contains(ugRow.Cells("ID").Value.ToString) Then
                                ' if any other row belonging to the updating row's parent
                                ' has "No" set to true, set delete variables to false
                                If ugRow.Cells("Surface Sealed" + vbCrLf + "No").Value Or _
                                    ugRow.Cells("Well Caps" + vbCrLf + "No").Value Then

                                    If Not slCCAT.Contains(ugRow.Cells("Well#").Value) Then
                                        slCCAT.Add(ugRow.Cells("Well#").Value, ugRow.Cells("Well#").Text)
                                    End If

                                    'deleteCitation = False
                                    deleteDiscrep = False
                                End If
                            End If
                        Next

                        For i As Integer = 0 To slCCAT.Count - 1
                            strUpdatedCCAT += slCCAT.GetByIndex(i).ToString + ", "
                        Next
                        If strUpdatedCCAT <> String.Empty Then
                            strUpdatedCCAT = strUpdatedCCAT.Trim.TrimEnd(",")
                        End If
                    Else
                        deleteCitation = False
                        deleteDiscrep = False
                    End If ' If checkList.Citation <> -1 Or checkList.DiscrepText <> String.Empty Then

                    If deleteCitation Then
                        For Each citation As MUSTER.Info.InspectionCitationInfo In oInspection.InspectionInfo.CitationsCollection.Values
                            If citation.QuestionID = checkList.ID Then
                                If strUpdatedCCAT <> String.Empty Then
                                    'citation.Reset()
                                    citation.CCAT = strUpdatedCCAT
                                    citation.Deleted = False
                                Else
                                    citation.Reset()
                                    citation.CCAT = String.Empty
                                    citation.Deleted = True
                                End If
                                Exit For
                            End If
                        Next
                    End If ' If deleteCitation Then

                    If deleteDiscrep Then
                        For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.InspectionInfo.DiscrepsCollection.Values
                            If discrep.QuestionID = checkList.ID Then
                                discrep.Reset()
                                discrep.Deleted = True
                                Exit For
                            End If
                        Next
                    End If ' If deleteDiscrep Then
                End If ' If Not checkList Is Nothing Then
            Else ' If msgBoxResult = DialogResult.Yes Then
                e.Cancel = True
            End If ' If msgBoxResult = DialogResult.Yes Then
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugCP_BeforeRowsDeleted(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles ugCP.BeforeRowsDeleted
        'Dim cp As MUSTER.Info.InspectionCPReadingsInfo
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim msgBoxResult As Windows.Forms.DialogResult
        Try
            e.DisplayPromptMsg = False
            msgBoxResult = MessageBox.Show("You have selected " + _
                            e.Rows.Length.ToString + " " + _
                            IIf(e.Rows.Length > 1, "rows", "row") + " for deletion." + vbCrLf + _
                            "Choose Yes to delete the " + _
                            IIf(e.Rows.Length > 1, "rows", "row") + _
                            " or No to exit.", "Delete Rows", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If msgBoxResult = DialogResult.Yes Then
                Dim alRowIDs As New ArrayList
                For Each dr In e.Rows
                    If Not alRowIDs.Contains(dr.Cells("ID").Value.ToString) Then
                        alRowIDs.Add(dr.Cells("ID").Value.ToString)
                    End If
                    If oInspection.CheckListMaster.InspectionCPReadings.ID <> dr.Cells("ID").Value Then
                        oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, dr.Cells("ID").Value)
                    End If
                    'cp = oInspection.InspectionInfo.CPReadingsCollection.Item(dr.Cells("ID").Value.ToString)
                    oInspection.CheckListMaster.InspectionCPReadings.Deleted = True
                    If oInspection.CheckListMaster.InspectionCPReadings.ID < 0 Then
                        oInspection.InspectionInfo.CPReadingsCollection.Remove(oInspection.CheckListMaster.InspectionCPReadings.ID)
                    End If
                Next
                Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo = oInspection.InspectionInfo.ChecklistMasterCollection.Item(e.Rows(0).Cells("QUESTION_ID").Value)
                If Not checkList Is Nothing Then
                    ' if there are any citations / discreps - delete
                    ' citation exists only if citation is not equal to -1
                    ' discrep exists only if DiscrepText is not empty
                    Dim strUpdatedCCAT As String = ""
                    Dim deleteCitation As Boolean = True
                    Dim deleteDiscrep As Boolean = True
                    If checkList.Citation <> -1 Or checkList.DiscrepText <> String.Empty Then
                        Dim slCCAT As New SortedList
                        deleteCitation = True
                        deleteDiscrep = True

                        For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows(0).ParentCollection
                            ' current row is not one of the deleted rows
                            If Not alRowIDs.Contains(ugRow.Cells("ID").Value.ToString) Then
                                ' if any other row belonging to the updating row's parent
                                ' has "Fail" or "Incon" set to true, set delete variables to false
                                If ugRow.Cells("Fail").Value Or ugRow.Cells("Incon").Value Then
                                    If ugRow.Cells("Line#").Text.StartsWith("3.5.4.") Then
                                        'strUpdatedCCAT += "T" + ugRow.Cells("Tank#").Text + ", "
                                        If Not slCCAT.Contains(ugRow.Cells("Tank#").Value) Then
                                            slCCAT.Add(ugRow.Cells("Tank#").Value, "T" + ugRow.Cells("Tank#").Text)
                                        End If
                                    ElseIf ugRow.Cells("Line#").Text.StartsWith("3.6.3.") Then
                                        'strUpdatedCCAT += "P" + ugRow.Cells("Pipe#").Text + ", "
                                        If Not slCCAT.Contains(ugRow.Cells("Pipe#").Value) Then
                                            slCCAT.Add(ugRow.Cells("Pipe#").Value, "P" + ugRow.Cells("Pipe#").Text)
                                        End If
                                    ElseIf ugRow.Cells("Line#").Text.StartsWith("3.7.6.") Then
                                        'strUpdatedCCAT += "TP" + ugRow.Cells("Term#").Text + ", "
                                        If Not slCCAT.Contains(ugRow.Cells("Term#").Value) Then
                                            slCCAT.Add(ugRow.Cells("Term#").Value, "TP" + ugRow.Cells("Term#").Text)
                                        End If
                                    Else
                                        deleteCitation = False
                                    End If
                                    deleteDiscrep = False
                                End If
                            End If
                        Next

                        For i As Integer = 0 To slCCAT.Count - 1
                            strUpdatedCCAT += slCCAT.GetByIndex(i).ToString + ", "
                        Next
                        If strUpdatedCCAT <> String.Empty Then
                            strUpdatedCCAT = strUpdatedCCAT.Trim.TrimEnd(",")
                        End If
                    End If ' If checkList.Citation <> -1 Or checkList.DiscrepText <> String.Empty Then

                    If deleteCitation Then
                        For Each citation As MUSTER.Info.InspectionCitationInfo In oInspection.InspectionInfo.CitationsCollection.Values
                            If citation.QuestionID = checkList.ID Then
                                If oInspection.CheckListMaster.InspectionCitation.ID <> citation.ID Then
                                    oInspection.CheckListMaster.InspectionCitation.Retrieve(oInspection.InspectionInfo, citation.ID)
                                End If
                                If strUpdatedCCAT <> String.Empty Then
                                    'citation.Reset()
                                    oInspection.CheckListMaster.InspectionCitation.CCAT = strUpdatedCCAT
                                    oInspection.CheckListMaster.InspectionCitation.Deleted = False
                                Else
                                    oInspection.CheckListMaster.InspectionCitation.Reset()
                                    oInspection.CheckListMaster.InspectionCitation.CCAT = String.Empty
                                    oInspection.CheckListMaster.InspectionCitation.Deleted = True
                                End If
                                Exit For
                            End If
                        Next
                    End If ' If deleteCitation Then

                    If deleteDiscrep Then
                        For Each discrep As MUSTER.Info.InspectionDiscrepInfo In oInspection.InspectionInfo.DiscrepsCollection.Values
                            If discrep.QuestionID = checkList.ID Then
                                If oInspection.CheckListMaster.InspectionDiscrep.ID <> discrep.ID Then
                                    oInspection.CheckListMaster.InspectionDiscrep.Retrieve(oInspection.InspectionInfo, discrep.ID)
                                End If
                                oInspection.CheckListMaster.InspectionDiscrep.Reset()
                                oInspection.CheckListMaster.InspectionDiscrep.Deleted = True
                                Exit For
                            End If
                        Next
                    End If ' If deleteDiscrep Then
                End If ' If Not checkList Is Nothing Then
            Else ' If msgBoxResult = DialogResult.Yes Then
                e.Cancel = True
            End If ' If msgBoxResult = DialogResult.Yes Then
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnAddTankCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddTankCP.Click
        Try
            Cursor.Current = Cursors.AppStarting
            AddTankPipeTermCPReadings(TankPipeTermCPReading.Tank)
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAddPipeCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPipeCP.Click
        Try
            Cursor.Current = Cursors.AppStarting
            AddTankPipeTermCPReadings(TankPipeTermCPReading.Pipe)
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAddTermCP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddTermCP.Click
        Try
            Cursor.Current = Cursors.AppStarting
            AddTankPipeTermCPReadings(TankPipeTermCPReading.Term)
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub AddTankPipeTermCPReadings(ByVal cpReading As TankPipeTermCPReading)
        Dim cp As MUSTER.Info.InspectionCPReadingsInfo
        Dim parentLineNum, childLineNum As String
        Dim bolFindRow As Boolean = False
        Try
            Select Case cpReading
                Case TankPipeTermCPReading.Tank
                    parentLineNum = "3.5.4"
                    childLineNum = "3.5.4." + oInspection.CheckListMaster.MaxTankCPLineNum.ToString
                Case TankPipeTermCPReading.Pipe
                    parentLineNum = "3.6.3"
                    childLineNum = "3.6.3." + oInspection.CheckListMaster.MaxPipeCPLineNum.ToString
                Case TankPipeTermCPReading.Term
                    parentLineNum = "3.7.6"
                    childLineNum = "3.7.6." + oInspection.CheckListMaster.MaxTermCPLineNum.ToString
            End Select
            If ugCP.ActiveRow Is Nothing Then
                bolFindRow = True
            ElseIf ugCP.ActiveRow.Cells("Line#").Value <> parentLineNum Then
                bolFindRow = True
            End If
            If bolFindRow Then
                For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCP.Rows
                    If Not dr.Cells("Line#").Value Is Nothing Then
                        If dr.Cells("Line#").Value = parentLineNum Then
                            ugCP.ActiveRow = dr
                            Exit For
                        End If
                    End If
                Next
            End If
            Select Case cpReading
                Case TankPipeTermCPReading.Tank
                    ugCP.DisplayLayout.Bands(5).AddNew()
                Case TankPipeTermCPReading.Pipe
                    ugCP.DisplayLayout.Bands(9).AddNew()
                Case TankPipeTermCPReading.Term
                    ugCP.DisplayLayout.Bands(13).AddNew()
            End Select
            ' after addnew(), active row is the new row added
            Select Case parentLineNum 'ugCP.ActiveRow.ParentRow.Cells("Line#").Value
                Case "3.5.4"
                    cp = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            oInspection.ID, _
                            CType(ugCP.ActiveRow.ParentRow.Cells("QUESTION_ID").Value, Int64), _
                            0, _
                            0, _
                            12, _
                            String.Empty, _
                            String.Empty, _
                            String.Empty, _
                            String.Empty, _
                            -1, _
                            False, _
                            String.Empty, _
                            dtNullDate, _
                            String.Empty, _
                            dtNullDate, _
                            oInspection.CheckListMaster.MaxTankCPLineNum, _
                            False, _
                            False, _
                            -1, False, False)
                    oInspection.CheckListMaster.InspectionCPReadings.Add(cp)
                    oInspection.CheckListMaster.MaxTankCPLineNum += 1
                Case "3.6.3"
                    cp = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            oInspection.ID, _
                            CType(ugCP.ActiveRow.ParentRow.Cells("QUESTION_ID").Value, Int64), _
                            0, _
                            0, _
                            10, _
                            String.Empty, _
                            String.Empty, _
                            String.Empty, _
                            String.Empty, _
                            -1, _
                            False, _
                            String.Empty, _
                            dtNullDate, _
                            String.Empty, _
                            dtNullDate, _
                            oInspection.CheckListMaster.MaxPipeCPLineNum, _
                            False, _
                            False, _
                            -1, False, False)
                    oInspection.CheckListMaster.InspectionCPReadings.Add(cp)
                    oInspection.CheckListMaster.MaxPipeCPLineNum += 1
                Case "3.7.6"
                    cp = New MUSTER.Info.InspectionCPReadingsInfo(0, _
                            oInspection.ID, _
                            CType(ugCP.ActiveRow.ParentRow.Cells("QUESTION_ID").Value, Int64), _
                            0, _
                            0, _
                            10, _
                            String.Empty, _
                            String.Empty, _
                            String.Empty, _
                            String.Empty, _
                            -1, _
                            False, _
                            String.Empty, _
                            dtNullDate, _
                            String.Empty, _
                            dtNullDate, _
                            oInspection.CheckListMaster.MaxTermCPLineNum, _
                            False, _
                            False, _
                            -1, False, False)
                    oInspection.CheckListMaster.InspectionCPReadings.Add(cp)
                    oInspection.CheckListMaster.MaxTermCPLineNum += 1
            End Select
            ugCP.ActiveRow.Cells("Line#").Value = childLineNum
            ugCP.ActiveRow.Cells("ID").Value = cp.ID
            ugCP.ActiveRow.Cells("INSPECTION_ID").Value = cp.InspectionID
            ugCP.ActiveRow.Cells("QUESTION_ID").Value = cp.QuestionID
            ugCP.ActiveRow.Cells("PASSFAILINCON").Value = cp.PassFailIncon
            ugCP.ActiveRow.Cells("CITATION").Value = oInspection.InspectionInfo.ChecklistMasterCollection.Item(cp.QuestionID).CCAT
            ugCP.ActiveRow.Cells("LINE_NUMBER").Value = cp.LineNumber
            ugCP.ActiveRow.Cells("Pass").Value = False
            ugCP.ActiveRow.Cells("Fail").Value = False
            ugCP.ActiveRow.Cells("Incon").Value = False
            btnSave.Enabled = oInspection.IsDirty Or oInspection.CheckListMaster.colIsDirty
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Tank Pipe Term"
    Private Sub ugTanks_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTanks.CellChange
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If "MATERIALS".Equals(e.Cell.Column.Key.ToUpper) Or _
                "OPTIONS".Equals(e.Cell.Column.Key.ToUpper) Or _
                "CP TYPE".Equals(e.Cell.Column.Key.ToUpper) Or _
                "LEAK DETECTION".Equals(e.Cell.Column.Key.ToUpper) Or _
                "OVERFILL TYPE".Equals(e.Cell.Column.Key.ToUpper) Or _
                "CONTENTS".Equals(e.Cell.Column.Key.ToUpper) Or _
                "FUEL TYPE".Equals(e.Cell.Column.Key.ToUpper) Then

                Cursor.Current = Cursors.AppStarting
                tankpipetermcellbeingupdated = True
                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Cell.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Cell.Row.Cells("TANK_ID").Value.ToString)
                    'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                End If

                'If "INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                '    UIUtilsGen.FillDateobjectValues(oTank.DateInstalledTank, e.Cell.Text)
                '    tankCellUpdated = True
                '    tankDateCellUpdated = True
                '    If Date.Compare(oTank.DateInstalledTank, dtNullDate) = 0 Then
                '        e.Cell.Value = DBNull.Value
                '    Else
                '        e.Cell.Value = e.Cell.Text
                '    End If
                'ElseIf "PLACED IN SERVICE".Equals(e.Cell.Column.Key.ToUpper) Then
                '    UIUtilsGen.FillDateobjectValues(oTank.PlacedInServiceDate, e.Cell.Text)
                '    tankCellUpdated = True
                '    tankDateCellUpdated = True
                '    If Date.Compare(oTank.PlacedInServiceDate, dtNullDate) = 0 Then
                '        e.Cell.Value = DBNull.Value
                '    Else
                '        e.Cell.Value = e.Cell.Text
                '    End If
                'ElseIf "SIZE".Equals(e.Cell.Column.Key.ToUpper) Then
                '    If oTank.Compartments.ID <> oTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString Then
                '        oTank.Compartments.Retrieve(oTank.TankInfo, oTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString)
                '    End If
                '    If IsNumeric(e.Cell.EditorResolved.Value) Then
                '        oTank.Compartments.Capacity = e.Cell.EditorResolved.Value
                '    Else
                '        oTank.Compartments.Capacity = 0
                '    End If
                '    tankCellUpdated = True
                '    e.Cell.Value = oTank.Compartments.Capacity
                '    ugTanks.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                If "CONTENTS".Equals(e.Cell.Column.Key.ToUpper) Then
                    If oTank.Compartments.ID <> oTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString Then
                        oTank.Compartments.Retrieve(oTank.TankInfo, oTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString)
                    End If
                    oTank.Compartments.Substance = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    tankCellUpdated = True
                    e.Cell.Value = oTank.Compartments.Substance

                    ' #2889
                    If e.Cell.Text.IndexOf("Used Oil") > -1 Then
                        e.Cell.Row.Cells("OVERFILL TYPE").Hidden = True
                        If oTank.OverFillType <> 0 Then
                            oTank.OverFillType = 0
                            regenerateCheckListItems = True
                        End If
                    Else
                        If oTank.SmallDelivery Then
                            e.Cell.Row.Cells("OVERFILL TYPE").Hidden = True
                            If oTank.OverFillType <> 0 Then
                                oTank.OverFillType = 0
                                regenerateCheckListItems = True
                            End If
                        Else
                            e.Cell.Row.Cells("OVERFILL TYPE").Hidden = False
                            ' Over Fill
                            If vListOverFill.FindByDataValue(e.Cell.Row.Cells("OVERFILL TYPE").Value) Is Nothing Then
                                e.Cell.Row.Cells("OVERFILL TYPE").Value = DBNull.Value
                                If oTank.OverFillType <> 0 Then
                                    oTank.OverFillType = 0
                                    regenerateCheckListItems = True
                                End If
                            End If
                        End If
                    End If
                ElseIf "FUEL TYPE".Equals(e.Cell.Column.Key.ToUpper) Then
                    If oTank.Compartments.ID <> oTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString Then
                        oTank.Compartments.Retrieve(oTank.TankInfo, oTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString)
                    End If
                    oTank.Compartments.FuelTypeId = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    tankCellUpdated = True
                    e.Cell.Value = oTank.Compartments.FuelTypeId
                ElseIf "MATERIALS".Equals(e.Cell.Column.Key.ToUpper) Then
                    oTank.TankMatDesc = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    tankCellUpdated = True
                    e.Cell.Value = oTank.TankMatDesc
                ElseIf "OPTIONS".Equals(e.Cell.Column.Key.ToUpper) Then
                    oTank.TankModDesc = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    tankCellUpdated = True
                    e.Cell.Value = oTank.TankModDesc
                ElseIf "CP TYPE".Equals(e.Cell.Column.Key.ToUpper) Then
                    oTank.TankCPType = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    tankCellUpdated = True
                    e.Cell.Value = oTank.TankCPType
                ElseIf "LEAK DETECTION".Equals(e.Cell.Column.Key.ToUpper) Then
                    oTank.TankLD = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    tankCellUpdated = True
                    e.Cell.Value = oTank.TankLD
                ElseIf "OVERFILL TYPE".Equals(e.Cell.Column.Key.ToUpper) Then
                    oTank.OverFillType = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    tankCellUpdated = True
                    e.Cell.Value = oTank.OverFillType
                    'ElseIf "LINED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oTank.LinedInteriorInstallDate, e.Cell.Text)
                    '    tankCellUpdated = True
                    '    tankDateCellUpdated = True
                    '    If Date.Compare(oTank.LinedInteriorInstallDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    '    If Not oInspection.CAPDatesEntered Then
                    '        oInspection.CAPDatesEntered = True
                    '    End If
                    'ElseIf "LINING INSPECTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oTank.LinedInteriorInspectDate, e.Cell.Text)
                    '    tankCellUpdated = True
                    '    tankDateCellUpdated = True
                    '    If Date.Compare(oTank.LinedInteriorInspectDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    '    If Not oInspection.CAPDatesEntered Then
                    '        oInspection.CAPDatesEntered = True
                    '    End If
                    'ElseIf "PTT".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oTank.TTTDate, e.Cell.Text)
                    '    tankCellUpdated = True
                    '    tankDateCellUpdated = True
                    '    If Date.Compare(oTank.TTTDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    '    If Not oInspection.CAPDatesEntered Then
                    '        oInspection.CAPDatesEntered = True
                    '    End If
                    'ElseIf "CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oTank.TCPInstallDate, e.Cell.Text)
                    '    tankCellUpdated = True
                    '    tankDateCellUpdated = True
                    '    If Date.Compare(oTank.TCPInstallDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    'ElseIf "CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oTank.LastTCPDate, e.Cell.Text)
                    '    tankCellUpdated = True
                    '    tankDateCellUpdated = True
                    '    If Date.Compare(oTank.LastTCPDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    '    If Not oInspection.CAPDatesEntered Then
                    '        oInspection.CAPDatesEntered = True
                    '    End If
                    'ElseIf "LAST USED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oTank.DateLastUsed, e.Cell.Text)
                    '    tankCellUpdated = True
                    '    tankDateCellUpdated = True
                    '    If Date.Compare(oTank.DateLastUsed, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                End If

                If tankCellUpdated Then
                    e.Cell.Row.Tag = 1
                    'oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, , bolReadOnly)
                    SetupTankPipeTermRow(TankPipeTermGrid.Tank, e.Cell.Row, bolReadOnly, oTank)
                    If tankDateCellUpdated Then
                        tankDateCellUpdated = False
                    Else
                        ugPipes.Focus()
                        ugTanks.Focus()
                    End If
                End If ' If tankCellUpdated Then
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            tankpipetermcellbeingupdated = False
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub ugTanks_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ugTanks.KeyDown
        If bolLoading Then Exit Sub
        If tankpipetermcellbeingupdated Then Exit Sub
        Dim tankpipetermcellbeingupdatedLocal As Boolean = tankpipetermcellbeingupdated
        Dim tankCellUpdatedLocal As Boolean = False
        Try
            If e.KeyCode = Keys.Delete Then
                If Not ugTanks.ActiveCell Is Nothing Then
                    If "OVERFILL TYPE".Equals(ugTanks.ActiveCell.Column.Key.ToUpper) Then
                        Cursor.Current = Cursors.AppStarting
                        tankpipetermcellbeingupdated = True
                        oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                        If oTank.TankId <> CType(ugTanks.ActiveCell.Row.Cells("TANK_ID").Value, Integer) Then
                            oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(ugTanks.ActiveCell.Row.Cells("TANK_ID").Value.ToString)
                            'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                        End If
                        oTank.OverFillType = 0
                        ugTanks.ActiveCell.Value = 0
                        tankCellUpdatedLocal = True
                        tankCellUpdated = True
                    End If
                End If
                'ElseIf e.KeyCode = Keys.Escape Then
                '    ' reset tank
                '    If Not ugTanks.ActiveRow Is Nothing Then
                '        Dim bolProceed As Boolean = False
                '        If ugTanks.ActiveRow.Tag Is Nothing Then
                '            bolProceed = False
                '        ElseIf ugTanks.ActiveRow.Tag = 1 Or ugTanks.ActiveRow.DataChanged Then
                '            bolProceed = True
                '        End If
                '        If bolProceed Then
                '            Cursor.Current = Cursors.AppStarting
                '            tankpipetermcellbeingupdated = True
                '            oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                '            If oTank.TankId <> CType(ugTanks.ActiveRow.Cells("TANK_ID").Value, Integer) Then
                '                oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(ugTanks.ActiveCell.Row.Cells("TANK_ID").Value.ToString)
                '                'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                '            End If
                '            oTank.Reset()
                '            Dim bolOriginalValueIsNothing, bolValueIsNothing As Boolean
                '            For Each ugCell As Infragistics.Win.UltraWinGrid.UltraGridCell In ugTanks.ActiveRow.Cells
                '                If ugCell.OriginalValue Is Nothing Then
                '                    bolOriginalValueIsNothing = True
                '                ElseIf ugCell.OriginalValue Is DBNull.Value Then
                '                    bolOriginalValueIsNothing = True
                '                Else
                '                    bolOriginalValueIsNothing = False
                '                End If
                '                If ugCell.Value Is Nothing Then
                '                    bolValueIsNothing = True
                '                ElseIf ugCell.Value Is DBNull.Value Then
                '                    bolValueIsNothing = True
                '                Else
                '                    bolValueIsNothing = False
                '                End If
                '                If Not bolOriginalValueIsNothing And bolValueIsNothing Then
                '                    ugCell.Value = ugCell.OriginalValue
                '                ElseIf bolOriginalValueIsNothing And Not bolValueIsNothing Then
                '                    ugCell.Value = DBNull.Value
                '                ElseIf Not bolOriginalValueIsNothing And Not bolValueIsNothing Then
                '                    ugCell.Value = ugCell.OriginalValue
                '                End If
                '            Next
                '            SetupTankPipeTermRow(TankPipeTermGrid.Tank, ugTanks.ActiveRow, bolReadOnly, oTank)
                '            ugTanks.ActiveRow.Tag = 0
                '        End If
                '    End If
            End If
            If tankCellUpdatedLocal Then
                'oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, , bolReadOnly)
                SetupTankPipeTermRow(TankPipeTermGrid.Tank, ugTanks.ActiveCell.Row, bolReadOnly, oTank)
                ugPipes.Focus()
                ugTanks.Focus()
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            tankpipetermcellbeingupdated = tankpipetermcellbeingupdatedLocal
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub ugTanks_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTanks.AfterCellUpdate
        If tankpipetermcellbeingupdated Then Exit Sub
        If bolLoading Then Exit Sub
        If bolFromBtnTanksPipes Then Exit Sub
        Dim tankpipetermcellbeingupdatedLocal As Boolean = tankpipetermcellbeingupdated
        Try
            If "SIZE".Equals(e.Cell.Column.Key.ToUpper) Or _
                "INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "PLACED IN SERVICE".Equals(e.Cell.Column.Key.ToUpper) Or _
                "LINED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "LINING INSPECTED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "PTT".Equals(e.Cell.Column.Key.ToUpper) Or _
                "CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "LAST USED".Equals(e.Cell.Column.Key.ToUpper) Then

                Cursor.Current = Cursors.AppStarting
                'Dim bolIsFacCapCandidate As Boolean = oInspection.CheckListMaster.Owner.Facilities.CAPCandidate
                tankpipetermcellbeingupdated = True

                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Cell.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Cell.Row.Cells("TANK_ID").Value.ToString)
                End If

                If "SIZE".Equals(e.Cell.Column.Key.ToUpper) Then
                    Cursor.Current = Cursors.AppStarting
                    oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                    If oTank.Compartments.ID <> oTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString Then
                        oTank.Compartments.Retrieve(oTank.TankInfo, oTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString)
                    End If
                    If IsNumeric(e.Cell.EditorResolved.Value) Then
                        oTank.Compartments.Capacity = e.Cell.EditorResolved.Value
                    Else
                        oTank.Compartments.Capacity = 0
                    End If
                    tankSizeUpdated = True
                    'e.Cell.Value = oTank.Compartments.Capacity
                    'ElseIf tankCellUpdated Then
                    '    oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, , bolReadOnly)
                    '    SetupTankPipeTermRow(TankPipeTermGrid.Tank, e.Cell.Row, bolReadOnly, oTank)
                    '    If tankDateCellUpdated Then
                    '        tankDateCellUpdated = False
                    '    Else
                    '        ugPipes.Focus()
                    '        ugTanks.Focus()
                    '    End If
                ElseIf "INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oTank.DateInstalledTank, e.Cell.Text)
                    tankCellUpdated = True
                    tankDateCellUpdated = True
                    If Date.Compare(oTank.DateInstalledTank, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "PLACED IN SERVICE".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oTank.PlacedInServiceDate, e.Cell.Text)
                    tankCellUpdated = True
                    tankDateCellUpdated = True
                    If Date.Compare(oTank.PlacedInServiceDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "LINED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oTank.LinedInteriorInstallDate, e.Cell.Text)
                    tankCellUpdated = True
                    tankDateCellUpdated = True
                    If Date.Compare(oTank.LinedInteriorInstallDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                ElseIf "LINING INSPECTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oTank.LinedInteriorInspectDate, e.Cell.Text)
                    tankCellUpdated = True
                    tankDateCellUpdated = True
                    If Date.Compare(oTank.LinedInteriorInspectDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                ElseIf "PTT".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oTank.TTTDate, e.Cell.Text)
                    tankCellUpdated = True
                    tankDateCellUpdated = True
                    If Date.Compare(oTank.TTTDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                ElseIf "CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oTank.TCPInstallDate, e.Cell.Text)
                    tankCellUpdated = True
                    tankDateCellUpdated = True
                    If Date.Compare(oTank.TCPInstallDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oTank.LastTCPDate, e.Cell.Text)
                    tankCellUpdated = True
                    tankDateCellUpdated = True
                    If Date.Compare(oTank.LastTCPDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                ElseIf "LAST USED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oTank.DateLastUsed, e.Cell.Text)
                    tankCellUpdated = True
                    tankDateCellUpdated = True
                    If Date.Compare(oTank.DateLastUsed, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                End If

                If tankCellUpdated Or tankSizeUpdated Then
                    'oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, , bolReadOnly)
                    SetupTankPipeTermRow(TankPipeTermGrid.Tank, e.Cell.Row, bolReadOnly, oTank)
                    If tankDateCellUpdated Then
                        tankDateCellUpdated = False
                    Else
                        ugPipes.Focus()
                        ugTanks.Focus()
                    End If
                End If ' If tankCellUpdated Then

            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            tankpipetermcellbeingupdated = tankpipetermcellbeingupdatedLocal
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub ugTanks_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugTanks.BeforeRowUpdate
        If tankpipetermcellbeingupdated Then Exit Sub
        Dim bolCapDataSavedThisTime As Boolean = False
        Try
            Cursor.Current = Cursors.AppStarting
            ' need to validate only if any cell was modified
            If tankCellUpdated Or tankSizeUpdated Then
                tankCellUpdated = False
                tankSizeUpdated = False
                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Row.Cells("TANK_ID").Value.ToString)
                    'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                End If

                If Not oTank.ValidateData(False, "INSPECTION") Then
                    e.Cancel = True
                    e.Row.Selected = True
                    'ugPipes.ActiveRow.Selected = False
                    'ugTerminations.ActiveRow.Selected = False
                Else
                    ' capture cap data before inspection overrides
                    If Date.Compare(oInspection.ChecklistFirstSaved, dtNullDate) = 0 Then
                        oInspection.ChecklistFirstSaved = Now
                        oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, True)
                        bolCapDataSavedThisTime = True
                    End If
                    ' save tank
                    oTank.ModifiedBy = MusterContainer.AppUser.ID
                    oTank.Save(moduleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        e.Cancel = True
                        e.Row.Selected = True
                        If bolCapDataSavedThisTime Then
                            oInspection.ChecklistFirstSaved = dtNullDate
                            oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, True, True)
                        End If
                        Exit Sub
                    End If
                    ' save to archive table only if inspection is submitted
                    ' set boolean to save data to archive table. placing after save so that if there is any error, it elimates saving to archive table
                    If Date.Compare(oInspection.SubmittedDate, dtNullDate) <> 0 Then
                        bolSaveDataToArchiveTbls = True
                    End If
                    e.Row.Tag = 0
                    regenerateCheckListItems = True
                    'Dim errStrLocal As String = String.Empty
                    'ugTankValidation(errStrLocal, e.Row)
                    'If errStrLocal.Length > 0 Then
                    '    errStrLocal = "The following field(s) are required for Tank #: " + e.Row.Cells("TANK #").Value.ToString + vbCrLf + errStrLocal
                    '    MessageBox.Show(errStrLocal, "Inspection Checklist", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    '    e.Cancel = True
                    '    e.Row.Selected = True
                    '    'ugPipes.ActiveRow.Selected = False
                    '    'ugTerminations.ActiveRow.Selected = False
                    'Else
                    '    ' save tank
                    '    oTank.ModifiedBy = MusterContainer.AppUser.ID
                    '    oTank.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True)
                    '    If Not UIUtilsGen.HasRights(returnVal) Then
                    '        e.Cancel = True
                    '        e.Row.Selected = True
                    '        Exit Sub
                    '    End If
                    'End If
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTanks_BeforeRowDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugTanks.BeforeRowDeactivate
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If Not ugTanks.ActiveRow Is Nothing Then
                Dim ea As New Infragistics.Win.UltraWinGrid.CancelableRowEventArgs(ugTanks.ActiveRow)
                ugTanks_BeforeRowUpdate(sender, ea)
                If ea.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTanks_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugTanks.Leave
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If Not ugTanks.ActiveRow Is Nothing Then
                Dim ea As New Infragistics.Win.UltraWinGrid.CancelableRowEventArgs(ugTanks.ActiveRow)
                ugTanks_BeforeRowUpdate(sender, ea)
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugPipes_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugPipes.CellChange
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If "STATUS".Equals(e.Cell.Column.Key.ToUpper) Or _
                "TYPE OF SYSTEM".Equals(e.Cell.Column.Key.ToUpper) Or _
                "MATERIALS".Equals(e.Cell.Column.Key.ToUpper) Or _
                "OPTIONS".Equals(e.Cell.Column.Key.ToUpper) Or _
                "BRAND".Equals(e.Cell.Column.Key.ToUpper) Or _
                "CP TYPE".Equals(e.Cell.Column.Key.ToUpper) Or _
                "PRI. LEAK DETECTION".Equals(e.Cell.Column.Key.ToUpper) Or _
                "SEC. LEAK DETECTION".Equals(e.Cell.Column.Key.ToUpper) Then

                Cursor.Current = Cursors.AppStarting
                tankpipetermcellbeingupdated = True
                Dim id As String
                id = e.Cell.Row.Cells("TANK_ID").Value.ToString + "|" + _
                        e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + _
                        e.Cell.Row.Cells("PIPE_ID").Value.ToString
                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Cell.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Cell.Row.Cells("TANK_ID").Value.ToString)
                    'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                End If
                oPipe = oTank.Pipes
                If oPipe.ID <> id Then
                    oPipe.Pipe = oTank.TankInfo.pipesCollection.Item(id)
                    'oPipe.Retrieve(oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.TankInfo, id)
                End If
                'tankpipetermcellbeingupdated = True
                If "STATUS".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.PipeStatusDesc = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pipeCellUpdated = True
                    e.Cell.Value = oPipe.PipeStatusDesc
                    'ElseIf "INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    'oPipe.PipeInstallDate = e.Cell.Text
                    '    UIUtilsGen.FillDateobjectValues(oPipe.PipeInstallDate, e.Cell.Text)
                    '    pipeCellUpdated = True
                    '    pipeDateCellUpdated = True
                    '    If Date.Compare(oPipe.PipeInstallDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    'ElseIf "PLACED IN SERVICE".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oPipe.PlacedInServiceDate, e.Cell.Text)
                    '    pipeCellUpdated = True
                    '    pipeDateCellUpdated = True
                    '    If Date.Compare(oPipe.PlacedInServiceDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                ElseIf "TYPE OF SYSTEM".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.PipeTypeDesc = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pipeCellUpdated = True
                    e.Cell.Value = oPipe.PipeTypeDesc
                ElseIf "MATERIALS".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.PipeMatDesc = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pipeCellUpdated = True
                    e.Cell.Value = oPipe.PipeMatDesc
                ElseIf "OPTIONS".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.PipeModDesc = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pipeCellUpdated = True
                    e.Cell.Value = oPipe.PipeModDesc
                ElseIf "BRAND".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.PipeManufacturer = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pipeCellUpdated = True
                    e.Cell.Value = oPipe.PipeManufacturer
                ElseIf "CP TYPE".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.PipeCPType = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pipeCellUpdated = True
                    e.Cell.Value = oPipe.PipeCPType
                ElseIf "PRI. LEAK DETECTION".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.PipeLD = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pipeCellUpdated = True
                    e.Cell.Value = oPipe.PipeLD
                    If Not e.Cell.OriginalValue Is DBNull.Value Then
                        If e.Cell.OriginalValue = 246 Or e.Cell.OriginalValue = 243 Then ' electronic or continuous
                            e.Cell.Row.Cells("SEC. LEAK DETECTION").Value = 0
                        End If
                    End If
                ElseIf "SEC. LEAK DETECTION".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.ALLDType = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pipeCellUpdated = True
                    e.Cell.Value = oPipe.ALLDType
                    If Not e.Cell.OriginalValue Is DBNull.Value Then
                        If e.Cell.OriginalValue <> e.Cell.Value Then
                            If e.Cell.OriginalValue = 497 And e.Cell.Value = 498 Then ' electronic to continuous
                                e.Cell.Row.Cells("PRI. LEAK DETECTION").Value = 243 ' continuous
                                oPipe.PipeLD = 243
                            ElseIf e.Cell.OriginalValue = 498 And e.Cell.Value = 497 Then ' continuous to electronic
                                e.Cell.Row.Cells("PRI. LEAK DETECTION").Value = 246 ' electronic
                                oPipe.PipeLD = 246
                            ElseIf e.Cell.Value = 498 Then ' continuous
                                e.Cell.Row.Cells("PRI. LEAK DETECTION").Value = 243 ' continuous
                                oPipe.PipeLD = 243
                            ElseIf e.Cell.Value = 497 Then ' electronic
                                e.Cell.Row.Cells("PRI. LEAK DETECTION").Value = 246 ' electronic
                                oPipe.PipeLD = 246
                            Else
                                e.Cell.Row.Cells("PRI. LEAK DETECTION").Value = 0
                                oPipe.PipeLD = 0
                            End If
                        End If
                    End If
                    'ElseIf "ALLD TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oPipe.ALLDTestDate, e.Cell.Text)
                    '    pipeCellUpdated = True
                    '    pipeDateCellUpdated = True
                    '    If Date.Compare(oPipe.ALLDTestDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    '    If Not oInspection.CAPDatesEntered Then
                    '        oInspection.CAPDatesEntered = True
                    '    End If
                    'ElseIf "PTT".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oPipe.LTTDate, e.Cell.Text)
                    '    pipeCellUpdated = True
                    '    pipeDateCellUpdated = True
                    '    If Date.Compare(oPipe.LTTDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    '    If Not oInspection.CAPDatesEntered Then
                    '        oInspection.CAPDatesEntered = True
                    '    End If
                    'ElseIf "CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oPipe.PipeCPInstalledDate, e.Cell.Text)
                    '    pipeCellUpdated = True
                    '    pipeDateCellUpdated = True
                    '    If Date.Compare(oPipe.PipeCPInstalledDate, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    'ElseIf "CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oPipe.PipeCPTest, e.Cell.Text)
                    '    pipeCellUpdated = True
                    '    pipeDateCellUpdated = True
                    '    If Date.Compare(oPipe.PipeCPTest, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    '    If Not oInspection.CAPDatesEntered Then
                    '        oInspection.CAPDatesEntered = True
                    '    End If
                    'ElseIf "LAST USED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oPipe.DateLastUsed, e.Cell.Text)
                    '    pipeCellUpdated = True
                    '    pipeDateCellUpdated = True
                    '    If Date.Compare(oPipe.DateLastUsed, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                End If
                If pipeCellUpdated Then
                    'oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, oPipe.Pipe, bolReadOnly)
                    SetupTankPipeTermRow(TankPipeTermGrid.Pipe, e.Cell.Row, bolReadOnly, oTank, oPipe)
                    If pipeDateCellUpdated Then
                        pipeDateCellUpdated = False
                    Else
                        ugTanks.Focus()
                        ugPipes.Focus()
                    End If
                End If
                tankpipetermcellbeingupdated = False
                Cursor.Current = Cursors.Default

            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugPipes_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugPipes.AfterCellUpdate
        If tankpipetermcellbeingupdated Then Exit Sub
        If bolLoading Then Exit Sub
        If bolFromBtnTanksPipes Then Exit Sub
        Dim tankpipetermcellbeingupdatedLocal As Boolean = tankpipetermcellbeingupdated
        Try
            If "INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "PLACED IN SERVICE".Equals(e.Cell.Column.Key.ToUpper) Or _
                "ALLD TESTED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "PTT".Equals(e.Cell.Column.Key.ToUpper) Or _
                "CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "LAST USED".Equals(e.Cell.Column.Key.ToUpper) Then

                Cursor.Current = Cursors.AppStarting
                'Dim bolIsFacCapCandidate As Boolean = oInspection.CheckListMaster.Owner.Facilities.CAPCandidate
                tankpipetermcellbeingupdated = True

                Dim id As String
                id = e.Cell.Row.Cells("TANK_ID").Value.ToString + "|" + _
                        e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + _
                        e.Cell.Row.Cells("PIPE_ID").Value.ToString
                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Cell.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Cell.Row.Cells("TANK_ID").Value.ToString)
                    'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                End If
                oPipe = oTank.Pipes
                If oPipe.ID <> id Then
                    oPipe.Pipe = oTank.TankInfo.pipesCollection.Item(id)
                End If

                If "INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                    'oPipe.PipeInstallDate = e.Cell.Text
                    UIUtilsGen.FillDateobjectValues(oPipe.PipeInstallDate, e.Cell.Text)
                    pipeCellUpdated = True
                    pipeDateCellUpdated = True
                    If Date.Compare(oPipe.PipeInstallDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "PLACED IN SERVICE".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oPipe.PlacedInServiceDate, e.Cell.Text)
                    pipeCellUpdated = True
                    pipeDateCellUpdated = True
                    If Date.Compare(oPipe.PlacedInServiceDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "ALLD TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oPipe.ALLDTestDate, e.Cell.Text)
                    pipeCellUpdated = True
                    pipeDateCellUpdated = True
                    If Date.Compare(oPipe.ALLDTestDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                ElseIf "PTT".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oPipe.LTTDate, e.Cell.Text)
                    pipeCellUpdated = True
                    pipeDateCellUpdated = True
                    If Date.Compare(oPipe.LTTDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                ElseIf "CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oPipe.PipeCPInstalledDate, e.Cell.Text)
                    pipeCellUpdated = True
                    pipeDateCellUpdated = True
                    If Date.Compare(oPipe.PipeCPInstalledDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oPipe.PipeCPTest, e.Cell.Text)
                    pipeCellUpdated = True
                    pipeDateCellUpdated = True
                    If Date.Compare(oPipe.PipeCPTest, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                ElseIf "LAST USED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oPipe.DateLastUsed, e.Cell.Text)
                    pipeCellUpdated = True
                    pipeDateCellUpdated = True
                    If Date.Compare(oPipe.DateLastUsed, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                End If

                If pipeCellUpdated Then
                    SetupTankPipeTermRow(TankPipeTermGrid.Pipe, e.Cell.Row, bolReadOnly, oTank, oPipe)
                    If pipeDateCellUpdated Then
                        pipeDateCellUpdated = False
                    Else
                        ugTanks.Focus()
                        ugPipes.Focus()
                    End If
                End If

            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            tankpipetermcellbeingupdated = tankpipetermcellbeingupdatedLocal
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub ugPipes_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugPipes.BeforeRowUpdate
        If tankpipetermcellbeingupdated Then Exit Sub
        Dim bolCapDataSavedThisTime As Boolean = False
        Try
            Cursor.Current = Cursors.AppStarting
            If pipeCellUpdated Then
                pipeCellUpdated = False
                Dim id As String
                id = e.Row.Cells("TANK_ID").Value.ToString + "|" + _
                        e.Row.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + _
                        e.Row.Cells("PIPE_ID").Value.ToString
                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Row.Cells("TANK_ID").Value.ToString)
                    'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                End If
                oPipe = oTank.Pipes
                If oPipe.ID <> id Then
                    oPipe.Pipe = oTank.TankInfo.pipesCollection.Item(id)
                    'oPipe.Retrieve(oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.TankInfo, id)
                End If

                If Not oPipe.ValidateData(False, "INSPECTION") Then
                    e.Cancel = True
                    e.Row.Selected = True
                    'ugTanks.ActiveRow.Selected = False
                    'ugTerminations.ActiveRow.Selected = False
                Else
                    ' capture cap data before inspection overrides
                    If Date.Compare(oInspection.ChecklistFirstSaved, dtNullDate) = 0 Then
                        oInspection.ChecklistFirstSaved = Now
                        oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, True)
                        bolCapDataSavedThisTime = True
                    End If
                    'save pipe
                    If oPipe.PipeID <= 0 Then
                        oPipe.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oPipe.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oPipe.Save(moduleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        e.Cancel = True
                        e.Row.Selected = True
                        If bolCapDataSavedThisTime Then
                            oInspection.ChecklistFirstSaved = dtNullDate
                            oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, True, True)
                        End If
                        Exit Sub
                    End If
                    ' save to archive table only if inspection is submitted
                    ' set boolean to save data to archive table. placing after save so that if there is any error, it elimates saving to archive table
                    If Date.Compare(oInspection.SubmittedDate, dtNullDate) <> 0 Then
                        bolSaveDataToArchiveTbls = True
                    End If
                    regenerateCheckListItems = True
                    'Dim errStrLocal As String = String.Empty
                    'ugPipeValidation(errStrLocal, e.Row)
                    'If errStrLocal.Length > 0 Then
                    '    errStrLocal = "The following field(s) are required for Pipe #: " + e.Row.Cells("PIPE #").Value.ToString + vbCrLf + errStrLocal
                    '    MessageBox.Show(errStrLocal, "Inspection Checklist", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    '    e.Cancel = True
                    '    e.Row.Selected = True
                    '    'ugTanks.ActiveRow.Selected = False
                    '    'ugTerminations.ActiveRow.Selected = False
                    'Else
                    '    'save pipe
                    '    If oPipe.PipeID <= 0 Then
                    '        oPipe.CreatedBy = MusterContainer.AppUser.ID
                    '    Else
                    '        oPipe.ModifiedBy = MusterContainer.AppUser.ID
                    '    End If
                    '    oPipe.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                    '    If Not UIUtilsGen.HasRights(returnVal) Then
                    '        Exit Sub
                    '    End If
                    'End If
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugPipes_BeforeRowDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugPipes.BeforeRowDeactivate
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If Not ugPipes.ActiveRow Is Nothing Then
                Dim ea As New Infragistics.Win.UltraWinGrid.CancelableRowEventArgs(ugPipes.ActiveRow)
                ugPipes_BeforeRowUpdate(sender, ea)
                If ea.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugPipes_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugPipes.Leave
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If Not ugPipes.ActiveRow Is Nothing Then
                Dim ea As New Infragistics.Win.UltraWinGrid.CancelableRowEventArgs(ugPipes.ActiveRow)
                ugPipes_BeforeRowUpdate(sender, ea)
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugTerminations_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTerminations.CellChange
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If "SUMP@TANK".Equals(e.Cell.Column.Key.ToUpper) Or _
                "SUMP@DISP".Equals(e.Cell.Column.Key.ToUpper) Or _
                "TYPE TERM@TANK".Equals(e.Cell.Column.Key.ToUpper) Or _
                "TYPE TERM@DISP".Equals(e.Cell.Column.Key.ToUpper) Or _
                "TANK TERM CP".Equals(e.Cell.Column.Key.ToUpper) Or _
                "DISP. TERM CP".Equals(e.Cell.Column.Key.ToUpper) Then

                Cursor.Current = Cursors.AppStarting
                tankpipetermcellbeingupdated = True
                Dim id As String
                id = e.Cell.Row.Cells("TANK_ID").Value.ToString + "|" + _
                        e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + _
                        e.Cell.Row.Cells("PIPE_ID").Value.ToString
                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Cell.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Cell.Row.Cells("TANK_ID").Value.ToString)
                    'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                End If
                oPipe = oTank.Pipes
                If oPipe.ID <> id Then
                    oPipe.Pipe = oTank.TankInfo.pipesCollection.Item(id)
                    'oPipe.Retrieve(oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.TankInfo, id)
                End If
                'tankpipetermcellbeingupdated = True
                If "SUMP@TANK".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.ContainSumpTank = e.Cell.Text
                    termCellUpdated = True
                    e.Cell.Value = e.Cell.Text
                ElseIf "SUMP@DISP".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.ContainSumpDisp = e.Cell.Text
                    termCellUpdated = True
                    e.Cell.Value = e.Cell.Text
                ElseIf "TYPE TERM@TANK".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.TermTypeTank = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    termCellUpdated = True
                    e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                ElseIf "TYPE TERM@DISP".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.TermTypeDisp = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    termCellUpdated = True
                    e.Cell.Value = oPipe.TermTypeDisp
                ElseIf "TANK TERM CP".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.TermCPTypeTank = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    termCellUpdated = True
                    e.Cell.Value = oPipe.TermCPTypeTank
                ElseIf "DISP. TERM CP".Equals(e.Cell.Column.Key.ToUpper) Then
                    oPipe.TermCPTypeDisp = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    termCellUpdated = True
                    e.Cell.Value = oPipe.TermCPTypeDisp
                    'ElseIf "TERM CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    '    UIUtilsGen.FillDateobjectValues(oPipe.TermCPLastTested, e.Cell.Text)
                    '    termCellUpdated = True
                    '    termDateCellUpdated = True
                    '    If Date.Compare(oPipe.TermCPLastTested, dtNullDate) = 0 Then
                    '        e.Cell.Value = DBNull.Value
                    '    Else
                    '        e.Cell.Value = e.Cell.Text
                    '    End If
                    '    If Not oInspection.CAPDatesEntered Then
                    '        oInspection.CAPDatesEntered = True
                    '    End If
                End If
                If termCellUpdated Then
                    'oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, oPipe.Pipe, bolReadOnly)
                    SetupTankPipeTermRow(TankPipeTermGrid.Term, e.Cell.Row, bolReadOnly, oTank, oPipe)
                    If termDateCellUpdated Then
                        termDateCellUpdated = False
                    Else
                        ugTanks.Focus()
                        ugTerminations.Focus()
                    End If
                End If
                tankpipetermcellbeingupdated = False
                Cursor.Current = Cursors.Default

            End If

        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTerminations_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTerminations.AfterCellUpdate
        If tankpipetermcellbeingupdated Then Exit Sub
        If bolLoading Then Exit Sub
        If bolFromBtnTanksPipes Then Exit Sub
        Dim tankpipetermcellbeingupdatedLocal As Boolean = tankpipetermcellbeingupdated
        Try
            If "TERM CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Or _
                "TERM CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then

                'Dim bolIsFacCapCandidate As Boolean = oInspection.CheckListMaster.Owner.Facilities.CAPCandidate
                Cursor.Current = Cursors.AppStarting
                tankpipetermcellbeingupdated = True
                Dim id As String
                id = e.Cell.Row.Cells("TANK_ID").Value.ToString + "|" + _
                        e.Cell.Row.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + _
                        e.Cell.Row.Cells("PIPE_ID").Value.ToString
                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Cell.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Cell.Row.Cells("TANK_ID").Value.ToString)
                    'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
                End If
                oPipe = oTank.Pipes
                If oPipe.ID <> id Then
                    oPipe.Pipe = oTank.TankInfo.pipesCollection.Item(id)
                    'oPipe.Retrieve(oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.TankInfo, id)
                End If

                If "TERM CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oPipe.TermCPLastTested, e.Cell.Text)
                    termCellUpdated = True
                    termDateCellUpdated = True
                    If Date.Compare(oPipe.TermCPLastTested, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                ElseIf "TERM CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
                    UIUtilsGen.FillDateobjectValues(oPipe.TermCPInstalledDate, e.Cell.Text)
                    termCellUpdated = True
                    termDateCellUpdated = True
                    If Date.Compare(oPipe.TermCPInstalledDate, dtNullDate) = 0 Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    If Not oInspection.CAPDatesEntered Then
                        oInspection.CAPDatesEntered = True
                    End If
                End If

                If termCellUpdated Then
                    SetupTankPipeTermRow(TankPipeTermGrid.Term, e.Cell.Row, bolReadOnly, oTank, oPipe)
                    If termDateCellUpdated Then
                        termDateCellUpdated = False
                    Else
                        ugTanks.Focus()
                        ugTerminations.Focus()
                    End If
                End If

            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            tankpipetermcellbeingupdated = tankpipetermcellbeingupdatedLocal
            Cursor.Current = Cursors.Default
        End Try
    End Sub
    Private Sub ugTerminations_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugTerminations.BeforeRowUpdate
        If tankpipetermcellbeingupdated Then Exit Sub
        Dim bolCapDataSavedThisTime As Boolean = False
        Try
            Cursor.Current = Cursors.AppStarting
            If termCellUpdated Then
                termCellUpdated = False
                Dim id As String
                id = e.Row.Cells("TANK_ID").Value.ToString + "|" + _
                        e.Row.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + _
                        e.Row.Cells("PIPE_ID").Value.ToString
                oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
                If oTank.TankId <> CType(e.Row.Cells("TANK_ID").Value, Integer) Then
                    oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Row.Cells("TANK_ID").Value.ToString)
                    'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Row.Cells("TANK_ID").Value, Integer))
                End If
                oPipe = oTank.Pipes
                If oPipe.ID <> id Then
                    oPipe.Pipe = oTank.TankInfo.pipesCollection.Item(id)
                    'oPipe.Retrieve(oInspection.CheckListMaster.Owner.Facilities.FacilityTanks.TankInfo, id)
                End If

                If Not oPipe.ValidateData(False, "INSPECTION") Then
                    e.Cancel = True
                    e.Row.Selected = True
                    'ugTanks.ActiveRow.Selected = False
                    'ugPipes.ActiveRow.Selected = False
                Else
                    ' capture cap data before inspection overrides
                    If Date.Compare(oInspection.ChecklistFirstSaved, dtNullDate) = 0 Then
                        oInspection.ChecklistFirstSaved = Now
                        oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, True)
                        bolCapDataSavedThisTime = True
                    End If
                    'save pipe
                    If oPipe.PipeID <= 0 Then
                        oPipe.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oPipe.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oPipe.Save(moduleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        e.Cancel = True
                        e.Row.Selected = True
                        If bolCapDataSavedThisTime Then
                            oInspection.ChecklistFirstSaved = dtNullDate
                            oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, True, True)
                        End If
                        Exit Sub
                    End If
                    ' save to archive table only if inspection is submitted
                    ' set boolean to save data to archive table. placing after save so that if there is any error, it elimates saving to archive table
                    If Date.Compare(oInspection.SubmittedDate, dtNullDate) <> 0 Then
                        bolSaveDataToArchiveTbls = True
                    End If
                    regenerateCheckListItems = True
                    'Dim errStrLocal As String = String.Empty
                    'ugTermValidation(errStrLocal, e.Row)
                    'If errStrLocal.Length > 0 Then
                    '    errStrLocal = "The following field(s) are required for Pipe #: " + e.Row.Cells("PIPE #").Value.ToString + vbCrLf + errStrLocal
                    '    MessageBox.Show(errStrLocal, "Inspection Checklist", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    '    e.Cancel = True
                    '    e.Row.Selected = True
                    '    'ugTanks.ActiveRow.Selected = False
                    '    'ugPipes.ActiveRow.Selected = False
                    'Else
                    '    'save pipe
                    '    If oPipe.PipeID <= 0 Then
                    '        oPipe.CreatedBy = MusterContainer.AppUser.ID
                    '    Else
                    '        oPipe.ModifiedBy = MusterContainer.AppUser.ID
                    '    End If
                    '    oPipe.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                    '    If Not UIUtilsGen.HasRights(returnVal) Then
                    '        Exit Sub
                    '    End If
                    'End If
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTerminations_BeforeRowDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ugTerminations.BeforeRowDeactivate
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If Not ugTerminations.ActiveRow Is Nothing Then
                Dim ea As New Infragistics.Win.UltraWinGrid.CancelableRowEventArgs(ugTerminations.ActiveRow)
                ugTerminations_BeforeRowUpdate(sender, ea)
                If ea.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTerminations_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugTerminations.Leave
        If tankpipetermcellbeingupdated Then Exit Sub
        Try
            If Not ugTerminations.ActiveRow Is Nothing Then
                Dim ea As New Infragistics.Win.UltraWinGrid.CancelableRowEventArgs(ugTerminations.ActiveRow)
                ugTerminations_BeforeRowUpdate(sender, ea)
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub ugTanks_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTanks.CellChange
    '    If tankpipetermcellbeingupdated Then Exit Sub
    '    Try
    '        Cursor.Current = Cursors.AppStarting
    '        tankpipetermcellbeingupdated = True
    '        oTank = oInspection.CheckListMaster.Owner.Facilities.FacilityTanks
    '        If oTank.TankId <> CType(e.Cell.Row.Cells("TANK_ID").Value, Integer) Then
    '            oTank.TankInfo = oInspection.CheckListMaster.Owner.Facility.TankCollection.Item(e.Cell.Row.Cells("TANK_ID").Value.ToString)
    '            'oTank.Retrieve(oInspection.CheckListMaster.Owner.Facility, CType(e.Cell.Row.Cells("FACILITY_ID").Value, Integer), CType(e.Cell.Row.Cells("TANK_ID").Value, Integer))
    '        End If
    '        tankpipetermcellbeingupdated = True
    '        If "INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
    '            UIUtilsGen.FillDateobjectValues(oTank.DateInstalledTank, e.Cell.Text)
    '            e.Cell.Value = e.Cell.Text
    '            tankCellUpdated = True
    '        ElseIf "MATERIALS".Equals(e.Cell.Column.Key.ToUpper) Then
    '            oTank.TankMatDesc = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
    '            e.Cell.Value = oTank.TankMatDesc
    '            tankCellUpdated = True
    '        ElseIf "OPTIONS".Equals(e.Cell.Column.Key.ToUpper) Then
    '            oTank.TankModDesc = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
    '            e.Cell.Value = oTank.TankModDesc
    '            tankCellUpdated = True
    '        ElseIf "CP TYPE".Equals(e.Cell.Column.Key.ToUpper) Then
    '            oTank.TankCPType = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
    '            e.Cell.Value = oTank.TankCPType
    '            tankCellUpdated = True
    '        ElseIf "LEAK DETECTION".Equals(e.Cell.Column.Key.ToUpper) Then
    '            oTank.TankLD = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
    '            e.Cell.Value = oTank.TankLD
    '            tankCellUpdated = True
    '        ElseIf "OVERFILL TYPE".Equals(e.Cell.Column.Key.ToUpper) Then
    '            oTank.OverFillType = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
    '            e.Cell.Value = oTank.OverFillType
    '            tankCellUpdated = True
    '        ElseIf "LINED".Equals(e.Cell.Column.Key.ToUpper) Then
    '            UIUtilsGen.FillDateobjectValues(oTank.LinedInteriorInstallDate, e.Cell.Text)
    '            e.Cell.Value = e.Cell.Text
    '            tankCellUpdated = True
    '        ElseIf "LINING INSPECTED".Equals(e.Cell.Column.Key.ToUpper) Then
    '            UIUtilsGen.FillDateobjectValues(oTank.LinedInteriorInspectDate, e.Cell.Text)
    '            e.Cell.Value = e.Cell.Text
    '            tankCellUpdated = True
    '        ElseIf "PTT".Equals(e.Cell.Column.Key.ToUpper) Then
    '            UIUtilsGen.FillDateobjectValues(oTank.TTTDate, e.Cell.Text)
    '            e.Cell.Value = e.Cell.Text
    '            tankCellUpdated = True
    '        ElseIf "CP INSTALLED".Equals(e.Cell.Column.Key.ToUpper) Then
    '            UIUtilsGen.FillDateobjectValues(oTank.TCPInstallDate, e.Cell.Text)
    '            e.Cell.Value = e.Cell.Text
    '            tankCellUpdated = True
    '        ElseIf "CP TESTED".Equals(e.Cell.Column.Key.ToUpper) Then
    '            UIUtilsGen.FillDateobjectValues(oTank.LastTCPDate, e.Cell.Text)
    '            e.Cell.Value = e.Cell.Text
    '            tankCellUpdated = True
    '        End If
    '        tankpipetermcellbeingupdated = False
    '        If tankCellUpdated Then
    '            oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, , bolReadOnly)
    '            SetupTankPipeTerm(TankPipeTermGrid.Tank, bolReadOnly)
    '            'tankCellUpdated = True
    '            'tankpipetermcellbeingupdated = True
    '            'ugPipes.Focus()
    '            'ugTanks.Focus()
    '            'tankpipetermcellbeingupdated = False
    '        End If
    '        Cursor.Current = Cursors.Default
    '    Catch ex As Exception
    '        Cursor.Current = Cursors.Default
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugTanks_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles ugTanks.BeforeCellUpdate
    '    If tankpipetermcellbeingupdated Then Exit Sub
    '    Try
    '        If tankCellUpdated Then
    '            oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, , bolReadOnly)
    '            SetupTankPipeTerm(TankPipeTermGrid.Tank, bolReadOnly)
    '            'tankCellUpdated = True
    '            tankpipetermcellbeingupdated = True
    '            ugPipes.Focus()
    '            ugTanks.Focus()
    '            tankpipetermcellbeingupdated = False
    '        End If
    '    Catch ex As Exception
    '        Cursor.Current = Cursors.Default
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugTanks_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTanks.AfterCellUpdate
    '    If tankpipetermcellbeingupdated Then Exit Sub
    '    Try
    '        If tankCellUpdated Then
    '            oInspection.CheckListMaster.CheckTankPipeBelongToCL(oTank.TankInfo, , bolReadOnly)
    '            SetupTankPipeTerm(TankPipeTermGrid.Tank, bolReadOnly)
    '            'tankCellUpdated = True
    '            tankpipetermcellbeingupdated = True
    '            ugPipes.Focus()
    '            ugTanks.Focus()
    '            tankpipetermcellbeingupdated = False
    '        End If
    '    Catch ex As Exception
    '        Cursor.Current = Cursors.Default
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region
#Region "Inspector"
    Private Sub ugInspector_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugInspector.BeforeRowUpdate
        Dim id As Int64
        Dim staffID, inspID As Int64
        Dim timein, timeout As String
        Dim inspDT As Date
        Dim bolDeleted As Boolean
        Try
            Cursor.Current = Cursors.AppStarting
            id = e.Row.Cells("INS_DATES_ID").Value
            inspID = e.Row.Cells("INSPECTION_ID").Value
            staffID = IIf(e.Row.Cells("DEQ INSPECTOR").Value Is DBNull.Value, 0, e.Row.Cells("DEQ INSPECTOR").Value)
            inspDT = IIf(e.Row.Cells("DATE INSPECTED").Value Is DBNull.Value, dtNullDate, e.Row.Cells("DATE INSPECTED").Value)
            timein = IIf(e.Row.Cells("TIME IN").Value Is DBNull.Value, String.Empty, e.Row.Cells("TIME IN").Value)
            timeout = IIf(e.Row.Cells("TIME OUT").Value Is DBNull.Value, String.Empty, e.Row.Cells("TIME OUT").Value)
            bolDeleted = IIf(e.Row.Cells("DELETED").Value Is DBNull.Value, False, e.Row.Cells("DELETED").Value)
            If timeout <> String.Empty Then
                Dim dtTimeIn As DateTime = DateTime.Parse(timein)
                Dim dtTimeOut As DateTime = DateTime.Parse(timeout)
                If Date.Compare(timeout, timein) < 0 Then
                    MsgBox("Time out must be greater than Time In")
                    e.Cancel = True
                    Exit Sub
                End If
            End If

            oInspection.CheckListMaster.PutCLInspectionHistory(id, inspID, staffID, inspDT, timein, timeout, bolDeleted, moduleID, MusterContainer.AppUser.ID, returnVal, MusterContainer.AppUser.UserKey)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If e.Row.Cells("INS_DATES_ID").Value Is DBNull.Value And id > 0 Then
                e.Row.Cells("INS_DATES_ID").Value = id
            ElseIf e.Row.Cells("INS_DATES_ID").Value <= 0 And id > 0 Then
                e.Row.Cells("INS_DATES_ID").Value = id
            End If

            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugInspector_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugInspector.CellChange
        Try
            If "DEQ INSPECTOR".Equals(e.Cell.Column.Key) Then
                e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
            ElseIf "DATE INSPECTED".Equals(e.Cell.Column.Key) Then
                e.Cell.Value = e.Cell.Text
            ElseIf "TIME IN".Equals(e.Cell.Column.Key) Then
                e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
            ElseIf "TIME OUT".Equals(e.Cell.Column.Key) Then
                e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
            End If
            If e.Cell.Row.Cells("INS_DATES_ID").Value Is DBNull.Value Then
                e.Cell.Row.Cells("INS_DATES_ID").Value = 0
                e.Cell.Row.Cells("DELETED").Value = False
                e.Cell.Row.Cells("INSPECTION_ID").Value = oInspection.ID
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugInspector_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles ugInspector.BeforeRowsDeleted
        Dim id As Int64
        Dim staffID, inspID As Int64
        Dim timein, timeout As String
        Dim inspDT As Date
        Dim bolDeleted As Boolean
        Try
            If ugInspector.Rows.Count = e.Rows.Length Then
                MsgBox("Atleast one row must exist in checklist")
                e.Cancel = True
                Exit Sub
            Else
                e.DisplayPromptMsg = False
                If MessageBox.Show("You have selected " + e.Rows.Length.ToString + " row(s) for deletion." + vbCrLf + "Choose Yes to delete the rows or No to exit.", "Delete Rows", MsgBoxStyle.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                    e.Cancel = True
                    Exit Sub
                End If
            End If
            Cursor.Current = Cursors.AppStarting
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
                If Not ugRow.IsAddRow Then
                    id = ugRow.Cells("INS_DATES_ID").Value
                    inspID = ugRow.Cells("INSPECTION_ID").Value
                    staffID = IIf(ugRow.Cells("DEQ INSPECTOR").Value Is DBNull.Value, 0, ugRow.Cells("DEQ INSPECTOR").Value)
                    inspDT = IIf(ugRow.Cells("DATE INSPECTED").Value Is DBNull.Value, dtNullDate, ugRow.Cells("DATE INSPECTED").Value)
                    timein = IIf(ugRow.Cells("TIME IN").Value Is DBNull.Value, String.Empty, ugRow.Cells("TIME IN").Value)
                    timeout = IIf(ugRow.Cells("TIME OUT").Value Is DBNull.Value, String.Empty, ugRow.Cells("TIME OUT").Value)
                    bolDeleted = True

                    oInspection.CheckListMaster.PutCLInspectionHistory(id, inspID, staffID, inspDT, timein, timeout, bolDeleted, moduleID, MusterContainer.AppUser.ID, returnVal, MusterContainer.AppUser.UserKey)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If
            Next
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub mskTxtOwnerPhone_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtOwnerPhone.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oInspection.CheckListMaster.Owner.PhoneNumberOne, mskTxtOwnerPhone.FormattedText.Trim.ToString)
    End Sub
    Private Sub txtFacilityName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityName.TextChanged
        Try
            oInspection.CheckListMaster.Owner.Facilities.Name = txtFacilityName.Text.Trim
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddress.TextChanged
        If bolLoading Then Exit Sub
        If txtAddress.Tag > 0 Then
            oInspection.CheckListMaster.Owner.Facilities.AddressID = Integer.Parse(Trim(txtAddress.Tag))
        Else
            oInspection.CheckListMaster.Owner.Facilities.AddressID = 0
        End If
    End Sub
    Private Sub txtAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddress.DoubleClick
        Try
            AddressForm = New Address(UIUtilsGen.EntityTypes.Facility, oInspection.CheckListMaster.Owner.Facilities.FacilityAddresses, "Facility", oInspection.CheckListMaster.Owner.AddressId, moduleID)
            AddressForm.ShowFIPS = False
            AddressForm.ShowDialog()
            ' update txtfacilityaddress text
            oInspection.CheckListMaster.Owner.Facilities.FacilityAddresses.Retrieve(oInspection.CheckListMaster.Owner.Facilities.AddressID)
            txtAddress.Text = UIUtilsGen.FormatAddress(oInspection.CheckListMaster.Owner.Facilities.FacilityAddresses, True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtAddress_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                txtAddress_DoubleClick(sender, New System.EventArgs)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub txtAddress_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddress.Enter
        Try
            If txtAddress.Text = String.Empty Then
                txtAddress_DoubleClick(sender, e)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtLatitudeDegree_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLatitudeDegree.TextChanged
        Try
            UIUtilsGen.FillDoubleObjectValues(oInspection.CheckListMaster.Owner.Facilities.LatitudeDegree, txtLatitudeDegree.Text.Trim)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtLatitudeMin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLatitudeMin.TextChanged
        Try
            UIUtilsGen.FillDoubleObjectValues(oInspection.CheckListMaster.Owner.Facilities.LatitudeMinutes, txtLatitudeMin.Text.Trim)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtLatitudeSec_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLatitudeSec.TextChanged
        Try
            UIUtilsGen.FillDoubleObjectValues(oInspection.CheckListMaster.Owner.Facilities.LatitudeSeconds, txtLatitudeSec.Text.Trim)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtLongitudeDegree_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLongitudeDegree.TextChanged
        Try
            UIUtilsGen.FillDoubleObjectValues(oInspection.CheckListMaster.Owner.Facilities.LongitudeDegree, txtLongitudeDegree.Text.Trim)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtLongitudeMin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLongitudeMin.TextChanged
        Try
            UIUtilsGen.FillDoubleObjectValues(oInspection.CheckListMaster.Owner.Facilities.LongitudeMinutes, txtLongitudeMin.Text.Trim)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtLongitudeSec_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLongitudeSec.TextChanged
        Try
            UIUtilsGen.FillDoubleObjectValues(oInspection.CheckListMaster.Owner.Facilities.LongitudeSeconds, txtLongitudeSec.Text.Trim)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkCapCandidate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCapCandidate.CheckedChanged
        Try
            oInspection.CheckListMaster.Owner.Facility.CAPCandidate = chkCapCandidate.Checked
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtComments_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComments.TextChanged
        Try
            oInspection.CheckListMaster.InspectionComments.InsComments = txtComments.Text
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtOwnersRep_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnersRep.TextChanged
        Try
            oInspection.OwnersRep = txtOwnersRep.Text.Trim
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Dim strTankPipeTermErr As String = String.Empty
        'Dim strResponseErr As String = String.Empty
        Dim bolValidateTankPipe As Boolean = Not regenerateCheckListItems
        Try
            CheckTankPipeTermModified()
            If ValidateCheckList(bolValidateTankPipe) Then
                If ValidateLatLong() Then
                    oInspection.LetterGenerated = False
                    If Date.Compare(oInspection.ChecklistFirstSaved, dtNullDate) = 0 Then
                        oInspection.ChecklistFirstSaved = Now
                        oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, True)
                    End If
                    If oInspection.ID <= 0 Then
                        oInspection.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oInspection.ModifiedBy = MusterContainer.AppUser.ID
                    End If

                    If Not MusterContainer.AppUser.HEAD_CANDE Then
                        oInspection.CAEViewed = dtNullDate
                    End If

                    Dim Success As Boolean = True
                    Success = oInspection.Save(moduleID, MusterContainer.AppUser.UserKey, returnVal, , , MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    Me.Text = IIf(bolReadOnly, "View", "Edit") + " CheckList for Facility : " + oInspection.CheckListMaster.Owner.Facilities.Name + " (" + oInspection.CheckListMaster.Owner.Facilities.ID.ToString + ")"

                    If Success Then
                        ' save tank/pipe/term to archive table only if boolean is set to save from tank/pipe/term grid
                        If bolSaveDataToArchiveTbls Then
                            oInspection.PutInspectionArchive(oInspection.FacilityID, oInspection.InspectionInfo.ID, moduleID, MusterContainer.AppUser.UserKey, returnVal, False)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                bolSaveDataToArchiveTbls = False
                                Exit Sub
                            End If
                            bolSaveDataToArchiveTbls = False
                        End If

                        MsgBox("CheckList Saved")
                        Me.CallingForm.Tag = "1"
                        'Else
                        '    oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, True, True)
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            bolCancel = True
            If Not Me.CallingForm Is Nothing Then
                If Not Me.CallingForm.Tag = "1" Then
                    Me.CallingForm.Tag = "0"
                End If
            End If
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSubmitToCE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmitToCE.Click
        Dim strTankPipeTermErr As String = String.Empty
        Dim bolValidateTankPipe As Boolean = Not regenerateCheckListItems
        Try
            CheckTankPipeTermModified()
            If ValidateCheckList(bolValidateTankPipe, False) Then
                If ValidateLatLong() Then
                    ' save soc values to the object
                    btnSoc.PerformClick()
                    SaveSOCValuesToObject()

                    ' #3030 Save inspection submitted date only after inspection responses and all other objects are saved.
                    oInspection.CheckListMaster.Flush(moduleID, MusterContainer.AppUser.UserKey, MusterContainer.AppUser.ID, returnVal, MusterContainer.AppUser.ID)
                    ' saving data here cause if in case there is an error saving data, the checklist will not be submitted
                    oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, False)
                    ' save to archive table irrespective of data changed or not
                    oInspection.PutInspectionArchive(oInspection.FacilityID, oInspection.InspectionInfo.ID, moduleID, MusterContainer.AppUser.UserKey, returnVal, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        bolSaveDataToArchiveTbls = False
                        Exit Sub
                    End If
                    bolSaveDataToArchiveTbls = False

                    oInspection.SubmittedDate = Now.Date
                    If oInspection.ID <= 0 Then
                        oInspection.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oInspection.ModifiedBy = MusterContainer.AppUser.ID
                    End If

                    If oInspection.Save(moduleID, MusterContainer.AppUser.UserKey, returnVal, , , MusterContainer.AppUser.ID) Then
                        MsgBox("Submitted to C & E")
                        bolReadOnly = True
                        MakeFormReadOnly(bolReadOnly)
                        btnSubmitToCE.Enabled = False
                        btnUnsubmit.Enabled = True
                        Me.CallingForm.Tag = "1"
                        Me.Text = "View CheckList for Facility : " + oInspection.CheckListMaster.Owner.Facilities.Name + " (" + oInspection.CheckListMaster.Owner.Facilities.ID.ToString + ")"
                        regenerateCheckListItems = False
                        btnMaster.PerformClick()
                    End If
                End If
            End If
            'Cursor.Current = Cursors.AppStarting
            'btnSave.PerformClick()
            'ValidateTankPipeTerm(strTankPipeTermErr)
            'Cursor.Current = Cursors.Default
            'If strTankPipeTermErr.Length > 0 Then
            '    MsgBox(strTankPipeTermErr)
            'Else
            '    oInspection.SubmittedDate = Now.Date
            '    If oInspection.Save() Then
            '        MsgBox("Submitted to C & E")
            '        btnSubmitToCE.Enabled = False
            '        btnUnsubmit.Enabled = True
            '        Me.CallingForm.Tag = "1"
            '    End If
            'End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnUnsubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUnsubmit.Click
        Try
            If Date.Compare(oInspection.Completed, dtNullDate) = 0 Then
                oInspection.SubmittedDate = dtNullDate
                If oInspection.InspectionInfo.SOCsCollection.Count > 0 Then
                    For Each socInfo As MUSTER.Info.InspectionSOCInfo In oInspection.InspectionInfo.SOCsCollection.Values
                        If oInspection.CheckListMaster.InspectionSOC.ID <> socInfo.ID Then
                            oInspection.CheckListMaster.InspectionSOC.Retrieve(oInspection.InspectionInfo, socInfo.ID)
                        End If
                        oInspection.CheckListMaster.InspectionSOC.CAEOverride = False
                    Next
                End If

                If oInspection.ID <= 0 Then
                    oInspection.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oInspection.ModifiedBy = MusterContainer.AppUser.ID
                End If

                Dim Success As Boolean = False
                Success = oInspection.Save(moduleID, MusterContainer.AppUser.UserKey, returnVal, , , MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                If Success Then
                    oInspection.SaveCAPTankPipeDataBeforeAfterInspection(oInspection.ID, False, True)
                    MsgBox("Inspection Unsubmitted")
                    bolReadOnly = False
                    MakeFormReadOnly(bolReadOnly)
                    btnSubmitToCE.Enabled = True
                    btnUnsubmit.Enabled = False
                    Me.CallingForm.Tag = "1"
                    Me.Text = "Edit CheckList for Facility : " + oInspection.CheckListMaster.Owner.Facilities.Name + " (" + oInspection.CheckListMaster.Owner.Facilities.ID.ToString + ")"
                    regenerateCheckListItems = True
                    btnMaster.PerformClick()
                End If
            Else
                MsgBox("Cannot unsubmit Completed inspection")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        PrintCheckList()
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            If (Not bolReadOnly) And (Not bolPrint) And (Not bolCancel) Then
                If oInspection.CheckListMaster.colIsDirty() Then
                    Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                    If Results = MsgBoxResult.Yes Then
                        btnSave.PerformClick()
                    ElseIf Results = MsgBoxResult.Cancel Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            End If
            'oInspection.CheckListMaster.Owner.Remove(oInspection.CheckListMaster.Owner.ID)
            'oInspection.CheckListMaster = New MUSTER.BusinessLogic.pInspectionChecklistMaster
            'oInspection.Remove(oInspection.ID)
            'oInspection = New MUSTER.BusinessLogic.pInspection
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "External Events"
    Private Sub TankValidationErr(ByVal tnkID As Integer, ByVal strMessage As String) Handles oTank.evtTankValidationErr
        Try
            MessageBox.Show(strMessage, "Tank Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PipeValidationErr(ByVal strMessage As String) Handles oPipe.evtPipeErr
        Try
            MessageBox.Show(strMessage, "Pipe Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub InspectionChanged(ByVal bolValue As Boolean) Handles oInspection.evtInspectionChanged
        If Not bolValue Then
            bolValue = oInspection.CheckListMaster.colIsDirty Or oInspection.IsDirty
        End If
        btnSave.Enabled = bolValue
    End Sub
    Private Sub TankChanged(ByVal bolValue As Boolean) Handles oTank.evtTankChanged
        If Not bolValue Then
            bolValue = oInspection.CheckListMaster.colIsDirty Or oInspection.IsDirty
        End If
        btnSave.Enabled = bolValue
    End Sub
    Private Sub PipeChanged(ByVal bolValue As Boolean) Handles oPipe.evtPipeChanged
        If Not bolValue Then
            bolValue = oInspection.CheckListMaster.colIsDirty Or oInspection.IsDirty
        End If
        btnSave.Enabled = bolValue
    End Sub
    Private Sub ltrGen_CheckListProgress(ByVal percent As Single) Handles ltrGen.CheckListProgress
        Try
            If frmCLProgress.ProgressBarValue >= frmCLProgress.ProgressBarMax Then
                frmCLProgress.Close()
            Else
                frmCLProgress.ProgressBarValue = percent
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub AddressForm_NewAddressID(ByVal MyAddressID As Integer) Handles AddressForm.NewAddressID
        Try
            txtAddress.Tag = MyAddressID
            oInspection.CheckListMaster.Owner.Facilities.AddressID = MyAddressID
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Contact Management"
    Private Sub ugContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ModifyContact(ugContacts))
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkOwnerShowContactsforAllModules_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOwnerShowContactsforAllModules.CheckedChanged
        SetFilterSub()
    End Sub
    Private Sub SetFilterSub()
        Try
            'If Not dsContacts Is Nothing Then
            SetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub chkOwnerShowRelatedContacts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOwnerShowRelatedContacts.CheckedChanged
    '    SetFilterSub()
    'End Sub
    Private Sub chkOwnerShowActiveOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOwnerShowActiveOnly.CheckedChanged
        SetFilterSub()
    End Sub
    Private Sub SetFilter()
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strentities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Dim strEntityAssocIDs As String = String.Empty
        Try
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If chkOwnerShowActiveOnly.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If
            nEntityType = 9
            If chkOwnerShowContactsforAllModules.Checked Then
                ' User has the ability to view the contacts associated for the entity in other modules
                Dim strFilterForAllModules As String = MusterContainer.pConStruct.GetContactsForAllModules(oInspection.FacilityID.ToString)
                strEntityAssocIDs = strFilterForAllModules
                nEntityID = oInspection.OwnerID
                nModuleID = 0
            Else
                nEntityID = oInspection.OwnerID
                nModuleID = moduleID
            End If

            If chkOwnerShowRelatedContacts.Checked Then
                Dim ds As DataSet = oInspection.CheckListMaster.Owner.RunSQLQuery("SELECT DISTINCT FACILITY_ID FROM TBLREG_FACILITY WHERE DELETED = 0 AND OWNER_ID = " + oInspection.OwnerID.ToString)
                strentities = oInspection.FacilityID.ToString
                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        strentities = String.Empty
                        For Each dr As DataRow In ds.Tables(0).Rows
                            strentities += dr("FACILITY_ID").ToString + ","
                        Next
                        If strentities = String.Empty Then
                            strentities = oInspection.FacilityID.ToString
                        Else
                            strentities = strentities.Trim.TrimEnd(",")
                        End If
                    End If
                End If
                nRelatedEntityType = 6
            Else
                strentities = String.Empty
            End If
            dsContactsLocal = MusterContainer.pConStruct.GetFilteredContacts(nEntityID, nModuleID, strentities, bolActive, strEntityAssocIDs, nEntityType, nRelatedEntityType)
            ugContacts.DataSource = dsContactsLocal

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'strFilterString = String.Empty
            'If chkOwnerShowActiveOnly.Checked Then
            '    strFilterString = "ACTIVE = 1"
            'End If

            'If chkOwnerShowContactsforAllModules.Checked Then
            '    ' User has the ability to view the contacts associated for the entity in other modules
            '    If Not strFilterString Is String.Empty Then
            '        strFilterString += "AND ENTITYID = " + oInspection.OwnerID.ToString
            '    Else
            '        strFilterString += "ENTITYID = " + oInspection.OwnerID.ToString
            '    End If
            'Else
            '    If Not strFilterString Is String.Empty Then
            '        strFilterString += "AND MODULEID = 615 And ENTITYID = " + oInspection.OwnerID.ToString
            '    Else
            '        strFilterString += "ENTITYID = " + oInspection.OwnerID.ToString
            '    End If
            'End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugContacts.DataSource = dsContacts.Tables(0).DefaultView
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerAddSearchContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerAddSearchContact.Click
        Try
            objCntSearch = New ContactSearch(oInspection.OwnerID, 9, "Inspection", MusterContainer.pConStruct)
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer)
        Dim dsContactsLocal As DataSet
        Try
            'dsContacts = MusterContainer.pConStruct.GetAll()
            'dsContacts.Tables(0).DefaultView.RowFilter = "MODULEID = 615 And ENTITYID = " + oInspection.OwnerID.ToString
            dsContactsLocal = MusterContainer.pConStruct.GetFilteredContacts(EntityID, 615)
            ugGrid.DataSource = dsContactsLocal.Tables(0).DefaultView  'dsContacts.Tables(0).DefaultView


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

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ugGrid.DisplayLayout.Bands(0).Columns("Parent_Contact").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("Child_Contact").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ContactAssocID").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ModuleID").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("ContactType").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("VENDOR_NUMBER").Hidden = True
            ugGrid.DisplayLayout.Bands(0).Columns("DISPLAYAS").Hidden = True

            Me.chkOwnerShowActiveOnly.Checked = False
            Me.chkOwnerShowActiveOnly.Checked = True

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerModifyContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerModifyContact.Click
        Try
            ModifyContact(ugContacts)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
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
            If UCase(dr.Cells("Module").Value) <> UCase("CheckList") Then
                MsgBox(" Cannot modify " + dr.Cells("Module").Value.ToString + " Contacts in CheckList")
                Exit Function
            End If
            If IsNothing(ContactFrm) Then
                ContactFrm = New Contacts(Integer.Parse(dr.Cells("EntityID").Value), CInt(dr.Cells("EntityType").Value), dr.Cells("Module").Value, CInt(dr.Cells("ContactID").Value), dr, MusterContainer.pConStruct, "MODIFY")
            End If
            ContactFrm.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
    Private Sub btnOwnerDeleteContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerDeleteContact.Click
        Try
            DeleteContact(ugContacts, oInspection.FacilityID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer)
        Try
            If ugGrid.Rows.Count <= 0 Then Exit Function

            If ugGrid.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete.")
                Exit Function
            End If
            If (CInt(ugGrid.ActiveRow.Cells("EntityID").Value) <> oInspection.OwnerID) Or (CInt(ugGrid.ActiveRow.Cells("ModuleID").Value) <> 615) Then
                MsgBox("Selected contact is not associated with the current entity and cannot be deleted")
                Exit Function
            End If
            result = MessageBox.Show("Are you sure you wish to DELETE the record?", "MUSTER", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.No Then Exit Function

            MusterContainer.pConStruct.Remove(ugGrid.ActiveRow.Cells("EntityAssocID").Text, moduleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Function
            End If
            ugGrid.ActiveRow.Delete(False)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Function
    Private Sub btnOwnerAssociateContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerAssociateContact.Click
        Try
            AssociateContact(ugContacts, oInspection.FacilityID, 6)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer)
        Try
            If ugGrid.Rows.Count <= 0 Then Exit Function

            If ugGrid.ActiveRow Is Nothing Then
                MsgBox("Select row to Associate.")
                Exit Function
            End If
            Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow = ugGrid.ActiveRow
            If ((CInt(ugGrid.ActiveRow.Cells("EntityID").Value) = nEntityID) And (CInt(ugGrid.ActiveRow.Cells("ModuleID").Value) = 615)) Then
                MsgBox("Selected contact is already associated with the current entity")
                Exit Function
            End If
            If IsNothing(ContactFrm) Then
                ContactFrm = New Contacts(oInspection.OwnerID, 9, "Inspection", CInt(dr.Cells("ContactID").Value), dr, MusterContainer.pConStruct, "ASSOCIATE")
            End If
            ContactFrm.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Sub Search_ContactAdded() Handles objCntSearch.ContactAdded
        LoadContacts(ugContacts, oInspection.OwnerID, 9)
        SetFilter()
    End Sub
    Private Sub Contact_ContactAdded() Handles ContactFrm.ContactAdded
        LoadContacts(ugContacts, oInspection.OwnerID, 9)
        SetFilter()
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
#Region "Flags"
    Private Sub btnFlags_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFlags.Click
        Try
            SF = New ShowFlags(oInspection.ID, UIUtilsGen.EntityTypes.Inspection, "Inspection")
            SF.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Function GetFlagsForPrintedChecklist() As String
        Dim strReturnVal As String = String.Empty
        If btnFacFlag.BackColor.Equals(Color.Red) Or btnFacFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "REGISTRATION" + ", "
        End If
        If btnFeeFlag.BackColor.Equals(Color.Red) Or btnFeeFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "FEES" + ", "
        End If
        If btnCloFlag.BackColor.Equals(Color.Red) Or btnCloFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "CLOSURE" + ", "
        End If
        If btnLUSFlag.BackColor.Equals(Color.Red) Or btnLUSFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "LUST" + ", "
        End If
        If btnFinFlag.BackColor.Equals(Color.Red) Or btnFinFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "FINANCIAL" + ", "
        End If
        If btnInsFlag.BackColor.Equals(Color.Red) Or btnInsFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "INSPECTION" + ", "
        End If
        If btnCandEFlag.BackColor.Equals(Color.Red) Or btnCandEFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "C && E" + ", "
        End If
        If btnFirmLicFlag.BackColor.Equals(Color.Red) Or btnFirmLicFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "COMPANY" + ", "
        End If
        If btnIndvLicFlag.BackColor.Equals(Color.Red) Or btnIndvLicFlag.BackColor.Equals(Color.Yellow) Then
            strReturnVal += "LICENSEE" + ", "
        End If
        If strReturnVal <> String.Empty Then
            strReturnVal = strReturnVal.Trim.TrimEnd(",")
        End If
        Return strReturnVal
    End Function
    Private Sub LoadBarometer()
        Try
            Dim ds As DataSet
            ds = pFlag.GetBarometerColors(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility)

            ClearBarometer()

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

            btnOwnerFlag.Enabled = True
            btnFacFlag.Enabled = True
            btnFeeFlag.Enabled = True
            btnCloFlag.Enabled = True
            btnLUSFlag.Enabled = True
            btnFinFlag.Enabled = True
            btnCandEFlag.Enabled = True
            btnInsFlag.Enabled = True
            btnFirmLicFlag.Enabled = True
            btnIndvLicFlag.Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ShowFlags(ByVal btn As Button)
        Try
            Select Case btn.Name
                Case btnOwnerFlag.Name
                    SF = New ShowFlags(oInspection.OwnerID, UIUtilsGen.EntityTypes.Owner, "Inspection")
                Case btnFacFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "Registration", True)
                Case btnFeeFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "Fees", True)
                Case btnCloFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "Closure", True)
                Case btnLUSFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "Technical", True)
                Case btnFinFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "Financial", True)
                Case btnCandEFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "C & E", True)
                Case btnInsFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "Inspection", True)
                Case btnFirmLicFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "Company", True)
                Case btnIndvLicFlag.Name
                    SF = New ShowFlags(oInspection.FacilityID, UIUtilsGen.EntityTypes.Facility, "Licensee", True)
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

    Private Sub SF_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles SF.FlagAdded
        LoadBarometer()
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        If Not Me.CallingForm Is Nothing Then
            Dim mc As MusterContainer = CallingForm.MdiParent
            If Not mc Is Nothing Then
                mc.RefreshCalendarInfo()
                mc.LoadDueToMeCalendar()
                mc.LoadToDoCalendar()
            End If
        End If
    End Sub

    Private Sub btnOwnerFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerFlag.Click
        Try
            'SF = New ShowFlags(oInspection.OwnerID, uiutilsgen.EntityTypes.Owner, "Inspection")
            'SF.ShowDialog()
            ShowFlags(btnOwnerFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFacFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacFlag.Click
        Try
            'SF = New ShowFlags(oInspection.FacilityID, uiutilsgen.EntityTypes.Facility, "Registration", True)
            'SF.ShowDialog()
            ShowFlags(btnFacFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFeeFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFeeFlag.Click
        Try
            'SF = New ShowFlags(oInspection.FacilityID, uiutilsgen.EntityTypes.Facility, "Fees", True)
            'SF.ShowDialog()
            ShowFlags(btnFeeFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCloFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCloFlag.Click
        Try
            'SF = New ShowFlags(oInspection.FacilityID, uiutilsgen.EntityTypes.Facility, "Closure", True)
            'SF.ShowDialog()
            ShowFlags(btnCloFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLUSFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLUSFlag.Click
        Try
            'SF = New ShowFlags(oInspection.FacilityID, uiutilsgen.EntityTypes.Facility, "Technical", True)
            'SF.ShowDialog()
            ShowFlags(btnLUSFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFinFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinFlag.Click
        Try
            'SF = New ShowFlags(oInspection.FacilityID, uiutilsgen.EntityTypes.Facility, "Financial", True)
            'SF.ShowDialog()
            ShowFlags(btnFinFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCandEFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCandEFlag.Click
        Try
            'SF = New ShowFlags(oInspection.FacilityID, uiutilsgen.EntityTypes.Facility, "C & E", True)
            'SF.ShowDialog()
            ShowFlags(btnCandEFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnInsFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsFlag.Click
        Try
            'SF = New ShowFlags(oInspection.ID, uiutilsgen.EntityTypes.Inspection, "Inspection", True)
            'SF.ShowDialog()
            ShowFlags(btnInsFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFirmLicFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirmLicFlag.Click
        Try
            'SF = New ShowFlags(oInspection.FacilityID, uiutilsgen.EntityTypes.Facility, "Company", True)
            'SF.ShowDialog()
            ShowFlags(btnFirmLicFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnIndvLicFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIndvLicFlag.Click
        Try
            'SF = New ShowFlags(oInspection.FacilityID, uiutilsgen.EntityTypes.Facility, "Licensee", True)
            'SF.ShowDialog()
            ShowFlags(btnIndvLicFlag)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
End Class

#Region "Remove Pencil Class"
' A Class that Implements a DrawFilter, 
' to Remove the pencil image that appears after editing a cell
Public Class Remove_Pencil
    Implements Infragistics.Win.IUIElementDrawFilter
    Public Function GetPhasesToFilter(ByRef drawParams As Infragistics.Win.UIElementDrawParams) As Infragistics.Win.DrawPhase Implements Infragistics.Win.IUIElementDrawFilter.GetPhasesToFilter
        ' Drawing RowSelector  Call DrawElement Before the Image is Drawn
        If TypeOf drawParams.Element Is Infragistics.Win.UltraWinGrid.RowSelectorUIElement Then
            Return Infragistics.Win.DrawPhase.BeforeDrawImage
        Else
            Return Infragistics.Win.DrawPhase.None
        End If
    End Function
    Public Function DrawElement(ByVal drawPhase As Infragistics.Win.DrawPhase, ByRef drawParams As Infragistics.Win.UIElementDrawParams) As Boolean Implements Infragistics.Win.IUIElementDrawFilter.DrawElement
        ' If the image isn't drawn yet, and the UIElement is a RowSelector
        If drawPhase = drawPhase.BeforeDrawImage And TypeOf drawParams.Element Is Infragistics.Win.UltraWinGrid.RowSelectorUIElement Then
            ' Get a handle of the row that is being drawn
            Dim row As Infragistics.Win.UltraWinGrid.UltraGridRow = CType(drawParams.Element.GetContext(Type.GetType("UltraGridRow")), Infragistics.Win.UltraWinGrid.UltraGridRow)
            ' If the row being draw is the active row
            If row.IsActiveRow Then
                ' Draw an arrow
                drawParams.DrawArrowIndicator(ScrollButton.Right, drawParams.Element.Rect, Infragistics.Win.UIElementButtonState.Indeterminate)
                ' Return true, to stop any other image from being drawn
                Return True
            Else
                Return True
            End If
        End If
        ' Else return false, to draw as normal
        Return False
    End Function
End Class
#End Region
