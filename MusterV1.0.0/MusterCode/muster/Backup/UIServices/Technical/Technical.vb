Imports System.Diagnostics
Imports System.Data.SqlClient
Imports System.Text
Public Class Technical
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.Technical.vb
    '   Provides the child MDI window for the technical module.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      12/??/04    Original class definition.
    '  1.1        JC      1/02/04     Altered popFacility to cast FacID to INTEGER
    '                                    before calling Retrieve method of facility
    '                                    so that the proper retrieve method is used.
    '  1.2        EN      02/20/05     Added Code to handle save and cancel button and bolloading 
    '  1.3        JC      08/02/05     Moved all comments related activity to Comments
    '                                   module and removed internal grid for comments.
    '-------------------------------------------------------------------------------
    '
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"
    Private WithEvents WordApp As Word.Application
    Dim dCol1 As New DataColumn
    Dim dCol2 As New DataColumn
    Dim dCol3 As New DataColumn
    Dim dCol4 As New DataColumn
    Dim dCol5 As New DataColumn
    Dim dCol6 As New DataColumn
    Dim dCol7 As New DataColumn

    Private strTFChecklistPath As String

    Friend dTableComments As New System.Data.DataTable("tComments")
    Friend dTableGlobal As New System.Data.DataTable("tGlobalComments")
    Friend dsComments As New System.Data.DataSet
    Private WithEvents pCommentLocal As New MUSTER.BusinessLogic.pComments

    Friend tblComments As DataTable

    Private WithEvents oOwner As MUSTER.BusinessLogic.pOwner
    Private WithEvents oLustEvent As New MUSTER.BusinessLogic.pLustEvent
    Private oFacility As MUSTER.BusinessLogic.pFacility
    Private oProps As New MUSTER.BusinessLogic.pPropertyType
    Private oAddressInfo As MUSTER.Info.AddressInfo
    Private dtNullDate As Date = CDate("01/01/0001")
    Public MyGuid As New System.Guid

    Public strFacilityIdTags As String
    Private strFacilityAddress As String
    Private bolLoading As Boolean = False
    Private bolProcessingCheck As Boolean = False
    Private bolTabClick As Boolean = False
    Private bolPMHead As Boolean = False
    Private PMHead_UserID As String = String.Empty
    Private PMHead_StaffID As Int16 = 0
    Private UserID As String = String.Empty
    Private UserStaffID As Int16 = 0
    Private nCurrentFacility As Int64 = 0
    Private nOriginalPM As Integer
    '  Private WithEvents ContactFrm As Contacts
    Private WithEvents objCntSearch As ContactSearch
    Dim dsContacts As DataSet
    Dim result As DialogResult
    Dim ttHistory As New ToolTip
    Private WithEvents SF As ShowFlags
    'Private oEntity As New MUSTER.BusinessLogic.pEntity
    Private bolFrmActivated As Boolean = False
    Dim nSavedEventID As Int64
    Public nLastEventID As Int64 = 0
    Public nFacilityID As Integer
    Private strLustEventIdTags As String
    Private oFinancialEvent As New MUSTER.BusinessLogic.pFinancial
    Private returnVal As String = String.Empty
    Private pConStruct As New MUSTER.BusinessLogic.pContactStruct
    Private bolMGPTFCheckList As Boolean = False
    Private bolAllowDocPop As Boolean
    Dim oLetter As New Reg_Letters
    Dim bolNewRelease As Boolean = False
    Friend nCurrentEventID As Integer = -1

#End Region

#Region "Company events and variables"
    Private WithEvents oCompanySearch As CompanySearch
    Dim pCompany As New MUSTER.BusinessLogic.pCompany
    Dim pLicensee As New MUSTER.BusinessLogic.pLicensee
    Dim strFromEracIrac As String
    #End Region

#Region " Windows Form Designer generated code "
    Public Sub New(ByVal OwnerID As Int64, ByVal FacilityId As Int64, ByRef POwn As MUSTER.BusinessLogic.pOwner, Optional ByVal nEventID As Integer = 0)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        InitializeComponent()

        If Not POwn Is Nothing Then
            oOwner = POwn
        Else
            oOwner = New MUSTER.BusinessLogic.pOwner
        End If
        oFacility = New MUSTER.BusinessLogic.pFacility
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Technical")
        InitControls()

        Dim oUser As New MUSTER.BusinessLogic.pUser
        Dim oUserInfo As New MUSTER.Info.UserInfo
        oUserInfo = oUser.RetrievePMHead
        PMHead_UserID = oUserInfo.ID
        PMHead_StaffID = oUserInfo.UserKey
        UserID = MusterContainer.AppUser.ID
        UserStaffID = MusterContainer.AppUser.UserKey
        If UserID.ToUpper = PMHead_UserID.ToUpper Then
            bolPMHead = True
        Else
            bolPMHead = False
        End If

        'Owner Info 
        If OwnerID > 0 Then
            PopulateOwnerInfo(OwnerID)
            If FacilityId <= 0 Then
                tbCntrlTechnical.SelectedTab = tbPageOwnerDetail
                Me.Text = "Technical - Owner Detail (" & txtOwnerName.Text & ")"
            End If
        End If
        If FacilityId > 0 Then 'Facility info .. 
            tbCntrlTechnical.SelectedTab = tbPageFacilityDetail
            PopFacility(FacilityId)
            CheckForSingleEvent_ByFacility(FacilityId)
        Else
            If OwnerID > 0 Then
                CheckForSingleEvent_ByOwner(OwnerID)
            End If
        End If
        If nEventID > 0 Then
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim sender As Object
            Dim e As System.EventArgs
            For Each ugRow In dgLUSTEvents.Rows
                If ugRow.Cells("EVENT_ID").Value = nEventID Then
                    dgLUSTEvents.ActiveRow = ugRow
                    Me.tbCntrlTechnical.SelectedTab = Me.tbPageLUSTEvent
                    nSavedEventID = 0
                    SetupModifyLustEventForm()

                    Exit Sub
                End If
            Next
        End If
    End Sub
    'Ends here ..
    Public Sub New(ByRef TheOwner As MUSTER.BusinessLogic.pOwner)
        MyBase.New()
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        oOwner = TheOwner
        'Added By Elango on Dec 27 2004 
        oFacility = New MUSTER.BusinessLogic.pFacility
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Technical")
        'Ends here 
        InitControls()
        PopulateOwnerInfo(TheOwner.ID)

        Dim oUser As New MUSTER.BusinessLogic.pUser
        Dim oUserInfo As New MUSTER.Info.UserInfo
        oUserInfo = oUser.RetrievePMHead
        PMHead_UserID = oUserInfo.ID
        PMHead_StaffID = oUserInfo.UserKey
        UserID = MusterContainer.AppUser.ID
        UserStaffID = MusterContainer.AppUser.UserKey
        If UserID.ToUpper = PMHead_UserID.ToUpper Then
            bolPMHead = True
        Else
            bolPMHead = False
        End If

        If TheOwner.ID > 0 Then
            CheckForSingleEvent_ByOwner(TheOwner.ID)
        End If

    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)

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
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents tbCntrlTechnical As System.Windows.Forms.TabControl
    Friend WithEvents tbPageOwnerDetail As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerBottom As System.Windows.Forms.Panel
    Friend WithEvents tbCtrlOwner As System.Windows.Forms.TabControl
    Friend WithEvents tbPageOwnerFacilities As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents lblNoOfFacilities As System.Windows.Forms.Label
    Public WithEvents chkOwnerAgencyInterest As System.Windows.Forms.CheckBox
    Public WithEvents lblOwnerActiveOrNot As System.Windows.Forms.Label
    Friend WithEvents LinkLblCAPSignup As System.Windows.Forms.LinkLabel
    Public WithEvents lblCAPParticipationLevel As System.Windows.Forms.Label
    Public WithEvents mskTxtOwnerFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone2 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone As AxMSMask.AxMaskEdBox
    Friend WithEvents lblOwnerEmail As System.Windows.Forms.Label
    Public WithEvents txtOwnerEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Public WithEvents txtOwnerAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerAddress As System.Windows.Forms.Label
    Public WithEvents txtOwnerName As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerStatus As System.Windows.Forms.Label
    Friend WithEvents lblOwnerCapParticipant As System.Windows.Forms.Label
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Friend WithEvents lblOwnerType As System.Windows.Forms.Label
    Public WithEvents txtOwnerAIID As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerAIID As System.Windows.Forms.Label
    Public WithEvents lblOwnerIDValue As System.Windows.Forms.Label
    Friend WithEvents lblOwnerID As System.Windows.Forms.Label
    Friend WithEvents lblOwnerPhone As System.Windows.Forms.Label
    Public WithEvents cmbOwnerType As System.Windows.Forms.ComboBox
    Friend WithEvents tbPageFacilityDetail As System.Windows.Forms.TabPage
    Friend WithEvents pnl_FacilityDetail As System.Windows.Forms.Panel
    Public WithEvents dtPickUpcomingInstallDateValue As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblUpcomingInstallDate As System.Windows.Forms.Label
    Public WithEvents chkUpcomingInstall As System.Windows.Forms.CheckBox
    Public WithEvents lblCAPStatusValue As System.Windows.Forms.Label
    Friend WithEvents lblCAPStatus As System.Windows.Forms.Label
    Public WithEvents txtFuelBrand As System.Windows.Forms.TextBox
    Friend WithEvents ll As System.Windows.Forms.Label
    Public WithEvents dtFacilityPowerOff As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblLUSTSite As System.Windows.Forms.Label
    Public WithEvents chkLUSTSite As System.Windows.Forms.CheckBox
    Friend WithEvents lblPowerOff As System.Windows.Forms.Label
    Friend WithEvents lblCAPCandidate As System.Windows.Forms.Label
    Public WithEvents chkCAPCandidate As System.Windows.Forms.CheckBox
    Friend WithEvents lblFacilityLocationType As System.Windows.Forms.Label
    Public WithEvents cmbFacilityLocationType As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacilityMethod As System.Windows.Forms.Label
    Public WithEvents cmbFacilityMethod As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacilityDatum As System.Windows.Forms.Label
    Public WithEvents cmbFacilityDatum As System.Windows.Forms.ComboBox
    Public WithEvents cmbFacilityType As System.Windows.Forms.ComboBox
    Public WithEvents txtFacilityLatSec As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLongSec As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatMin As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLongMin As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityLongMin As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongSec As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatMin As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatSec As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongDegree As System.Windows.Forms.Label
    Public WithEvents mskTxtFacilityFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtFacilityPhone As AxMSMask.AxMaskEdBox
    Public WithEvents txtFacilityAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilitySIC As System.Windows.Forms.Label
    Friend WithEvents lblFacilityFax As System.Windows.Forms.Label
    Public WithEvents dtPickFacilityRecvd As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateReceived As System.Windows.Forms.Label
    Public WithEvents chkSignatureofNF As System.Windows.Forms.CheckBox
    Friend WithEvents lblFacilitySigOnNF As System.Windows.Forms.Label
    Friend WithEvents lblFacilityFuelBrand As System.Windows.Forms.Label
    Public WithEvents txtFacilityLongDegree As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatDegree As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityLongitude As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatitude As System.Windows.Forms.Label
    Friend WithEvents lblFacilityType As System.Windows.Forms.Label
    Public WithEvents txtFacilityAIID As System.Windows.Forms.TextBox
    Friend WithEvents lblfacilityAIID As System.Windows.Forms.Label
    Public WithEvents lblFacilityIDValue As System.Windows.Forms.Label
    Friend WithEvents lblFacilityID As System.Windows.Forms.Label
    Friend WithEvents lblFacilityPhone As System.Windows.Forms.Label
    Public WithEvents txtFacilityName As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityAddress As System.Windows.Forms.Label
    Friend WithEvents lblFacilityName As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatDegree As System.Windows.Forms.Label
    Friend WithEvents tbPageLUSTEvent As System.Windows.Forms.TabPage
    Friend WithEvents tbPageSummary As System.Windows.Forms.TabPage
    Friend WithEvents Panel12 As System.Windows.Forms.Panel
    Friend WithEvents lnkLblNextFac As System.Windows.Forms.LinkLabel
    Friend WithEvents lnkLblPrevFacility As System.Windows.Forms.LinkLabel
    Friend WithEvents pnlFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents lblTotalNoOfLUSTEvents As System.Windows.Forms.Label
    Friend WithEvents btnAddLUSTEvent As System.Windows.Forms.Button
    Friend WithEvents btnTransferLUSTEvent As System.Windows.Forms.Button
    Friend WithEvents btnFacilityCancel As System.Windows.Forms.Button
    Friend WithEvents btnFacilitySave As System.Windows.Forms.Button
    Public WithEvents lblNoOfFacilitiesValue As System.Windows.Forms.Label
    Public WithEvents pnlOwnerDetail As System.Windows.Forms.Panel
    Public WithEvents dgLUSTEvents As Infragistics.Win.UltraWinGrid.UltraGrid
    Public WithEvents lblTotalNoOfLUSTEventsValue As System.Windows.Forms.Label
    Public WithEvents ugFacilityList As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlLUSTEventBottom As System.Windows.Forms.Panel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnDeleteLUSTEvent As System.Windows.Forms.Button
    Friend WithEvents btnSaveLUSTEvent As System.Windows.Forms.Button
    Public WithEvents lblFacilityStatusValue As System.Windows.Forms.Label
    Friend WithEvents lblFacilityStatus As System.Windows.Forms.Label
    Friend WithEvents pnlLUSTEventHeader As System.Windows.Forms.Panel
    Friend WithEvents lblEventIDValue As System.Windows.Forms.Label
    Friend WithEvents lblPrioritytt As System.Windows.Forms.Label
    Friend WithEvents cmbPriority As System.Windows.Forms.ComboBox
    Friend WithEvents lblPriority As System.Windows.Forms.Label
    Friend WithEvents cmbMGPTFStatus As System.Windows.Forms.ComboBox
    Friend WithEvents lblMGPTFStatus As System.Windows.Forms.Label
    Friend WithEvents cmbEventStatus As System.Windows.Forms.ComboBox
    Friend WithEvents lblEventStatus As System.Windows.Forms.Label
    Friend WithEvents lblEventCountValue As System.Windows.Forms.Label
    Public WithEvents cmbProjectManager As System.Windows.Forms.ComboBox
    Friend WithEvents lblEventID As System.Windows.Forms.Label
    Friend WithEvents pnlLustEvents As System.Windows.Forms.Panel
    Friend WithEvents pnlLUSTEventDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlEventInfo As System.Windows.Forms.Panel
    Friend WithEvents lblEventInfoHead As System.Windows.Forms.Label
    Friend WithEvents lblEventInfoDisplay As System.Windows.Forms.Label
    Friend WithEvents PnlEventInfoDetails As System.Windows.Forms.Panel
    Friend WithEvents dtLastGWS As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtLastPTT As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtLastLDR As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblLastGWS As System.Windows.Forms.Label
    Friend WithEvents txtRelatedSites As System.Windows.Forms.TextBox
    Friend WithEvents lblRelatedSites As System.Windows.Forms.Label
    Friend WithEvents cmbReleaseStatus As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSuspectedSource As System.Windows.Forms.ComboBox
    Friend WithEvents lblSuspectedSource As System.Windows.Forms.Label
    Friend WithEvents lblLastPTT As System.Windows.Forms.Label
    Friend WithEvents lblLastLDR As System.Windows.Forms.Label
    Friend WithEvents lblReleaseStatus As System.Windows.Forms.Label
    Friend WithEvents lblDateofReport As System.Windows.Forms.Label
    Friend WithEvents dtDateofReport As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtCompAssDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblEventStartDate As System.Windows.Forms.Label
    Friend WithEvents pnlReleaseInfo As System.Windows.Forms.Panel
    Friend WithEvents lblReleaseInfoHead As System.Windows.Forms.Label
    Friend WithEvents lblReleaseInfoDisplay As System.Windows.Forms.Label
    Friend WithEvents PnlReleaseInfoDetails As System.Windows.Forms.Panel
    Friend WithEvents btnEvtTankCollapse As System.Windows.Forms.Button
    Friend WithEvents btnEvtTankToggle As System.Windows.Forms.Button
    Friend WithEvents dtConfirmedOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents chkFreeProductUnKnown As System.Windows.Forms.CheckBox
    Friend WithEvents chkFreeProductWasteOil As System.Windows.Forms.CheckBox
    Friend WithEvents chkVaporPAH As System.Windows.Forms.CheckBox
    Friend WithEvents chkVaporBTEX As System.Windows.Forms.CheckBox
    Friend WithEvents chkToCVapor As System.Windows.Forms.CheckBox
    Friend WithEvents chkFreeProductKerosene As System.Windows.Forms.CheckBox
    Friend WithEvents chkFreeProductDiesel As System.Windows.Forms.CheckBox
    Friend WithEvents chkFreeProductGasoline As System.Windows.Forms.CheckBox
    Friend WithEvents chkToCFreeProduct As System.Windows.Forms.CheckBox
    Friend WithEvents chkGroundWaterTPH As System.Windows.Forms.CheckBox
    Friend WithEvents chkGroundWaterPAH As System.Windows.Forms.CheckBox
    Friend WithEvents chkGroundWaterBTEX As System.Windows.Forms.CheckBox
    Friend WithEvents chkGroundWater As System.Windows.Forms.CheckBox
    Friend WithEvents chkSoilTPH As System.Windows.Forms.CheckBox
    Friend WithEvents chkSoilPAH As System.Windows.Forms.CheckBox
    Friend WithEvents chkSoilBTEX As System.Windows.Forms.CheckBox
    Friend WithEvents chkSoil As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkInspection As System.Windows.Forms.CheckBox
    Friend WithEvents chkTankClosure As System.Windows.Forms.CheckBox
    Friend WithEvents chkInventoryShortage As System.Windows.Forms.CheckBox
    Friend WithEvents chkFailedPTT As System.Windows.Forms.CheckBox
    Friend WithEvents chkSoilContamination As System.Windows.Forms.CheckBox
    Friend WithEvents chkFreeProduct As System.Windows.Forms.CheckBox
    Friend WithEvents chkVapors As System.Windows.Forms.CheckBox
    Friend WithEvents chkGWContamination As System.Windows.Forms.CheckBox
    Friend WithEvents chkGWWell As System.Windows.Forms.CheckBox
    Friend WithEvents chkSurfaceSheen As System.Windows.Forms.CheckBox
    Friend WithEvents chkFacilityLeakDetection As System.Windows.Forms.CheckBox
    Friend WithEvents ugTankandPipes As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmbExtent As System.Windows.Forms.ComboBox
    Friend WithEvents lblExtent As System.Windows.Forms.Label
    Friend WithEvents cmbLocation As System.Windows.Forms.ComboBox
    Friend WithEvents cmbIdentifiedBy As System.Windows.Forms.ComboBox
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblIdentifiedBy As System.Windows.Forms.Label
    Friend WithEvents lblConfirmedOn As System.Windows.Forms.Label
    Friend WithEvents pnlFundsEligibility As System.Windows.Forms.Panel
    Friend WithEvents lblFundsEligibilityHead As System.Windows.Forms.Label
    Friend WithEvents lblFundsEligibilityDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlFundsEligibilityDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlFundsEligibilityQuestions As System.Windows.Forms.Panel
    Friend WithEvents chkQuestion17NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion17No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion17Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion17 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion16NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion16No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion16Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion16 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion15NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion15No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion15Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion15 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion14NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion14No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion14Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion14 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion10NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion10No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion10Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion10 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion8NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion8No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion8Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion8 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion6NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion6No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion6Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion6 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion5NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion5No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion5Yes As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion2No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion2Yes As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion1No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion1Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion5 As System.Windows.Forms.Label
    Friend WithEvents lblQuestion2 As System.Windows.Forms.Label
    Friend WithEvents lblQuestion1 As System.Windows.Forms.Label
    Friend WithEvents lblEligibilityComments As System.Windows.Forms.Label
    Friend WithEvents txtEligibilityComments As System.Windows.Forms.TextBox
    Friend WithEvents btnSendtoPM As System.Windows.Forms.Button
    Friend WithEvents btnViewCheckList As System.Windows.Forms.Button
    Friend WithEvents lblQuestionsNA As System.Windows.Forms.Label
    Friend WithEvents lblQuestionsNo As System.Windows.Forms.Label
    Friend WithEvents lblQuestionsYes As System.Windows.Forms.Label
    Friend WithEvents lblQuestions As System.Windows.Forms.Label
    Friend WithEvents pnlTFAssess As System.Windows.Forms.Panel
    Friend WithEvents dtCommissionOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtOPCHeadOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtUSTChiefOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPMHeadOn As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkforCommission As System.Windows.Forms.CheckBox
    Friend WithEvents chkforHeadofOPC As System.Windows.Forms.CheckBox
    Friend WithEvents lblCommissionBy As System.Windows.Forms.Label
    Friend WithEvents lblOPCHeadBy As System.Windows.Forms.Label
    Friend WithEvents lblUSTChiefBy As System.Windows.Forms.Label
    Friend WithEvents txtCommissionBy As System.Windows.Forms.TextBox
    Friend WithEvents txtOPCHeadBy As System.Windows.Forms.TextBox
    Friend WithEvents txtUSTChiefBy As System.Windows.Forms.TextBox
    Friend WithEvents txtPMHeadBy As System.Windows.Forms.TextBox
    Friend WithEvents lblPMHeadBy As System.Windows.Forms.Label
    Friend WithEvents lblCommissionOn As System.Windows.Forms.Label
    Friend WithEvents lblOPCHeadOn As System.Windows.Forms.Label
    Friend WithEvents lblUSTChiefOn As System.Windows.Forms.Label
    Friend WithEvents lblPMHeadOn As System.Windows.Forms.Label
    Friend WithEvents chkCommissionNo As System.Windows.Forms.CheckBox
    Friend WithEvents chkCommissionYes As System.Windows.Forms.CheckBox
    Friend WithEvents lblCommission As System.Windows.Forms.Label
    Friend WithEvents chkOPCHeadNo As System.Windows.Forms.CheckBox
    Friend WithEvents chkOPCHeadYes As System.Windows.Forms.CheckBox
    Friend WithEvents lblOPCHead As System.Windows.Forms.Label
    Friend WithEvents chkUSTChiefUndecided As System.Windows.Forms.CheckBox
    Friend WithEvents chkUSTChiefNo As System.Windows.Forms.CheckBox
    Friend WithEvents chkUSTChiefYes As System.Windows.Forms.CheckBox
    Friend WithEvents lblUSTChief As System.Windows.Forms.Label
    Friend WithEvents lblUndecided As System.Windows.Forms.Label
    Friend WithEvents lblNo As System.Windows.Forms.Label
    Friend WithEvents lblYes As System.Windows.Forms.Label
    Friend WithEvents chkPMHeadUndecided As System.Windows.Forms.CheckBox
    Friend WithEvents chkPMHeadYes As System.Windows.Forms.CheckBox
    Friend WithEvents lblPMHead As System.Windows.Forms.Label
    Friend WithEvents pnlActivitiesDocuments As System.Windows.Forms.Panel
    Friend WithEvents lblActivitiesDocumentsHead As System.Windows.Forms.Label
    Friend WithEvents lblActivitiesDocumentsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlActivitiesDocumentsDetails As System.Windows.Forms.Panel
    Friend WithEvents btnDeleteActivity As System.Windows.Forms.Button
    Friend WithEvents btnDeleteDocument As System.Windows.Forms.Button
    Friend WithEvents btnModifyDocument As System.Windows.Forms.Button
    Friend WithEvents btnAddDocument As System.Windows.Forms.Button
    Friend WithEvents btnModifyActivity As System.Windows.Forms.Button
    Friend WithEvents btnAddActivity As System.Windows.Forms.Button
    Friend WithEvents ugActivitiesandDocuments As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents chkShowallDocuments As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowallActivities As System.Windows.Forms.CheckBox
    Friend WithEvents pnlComments As System.Windows.Forms.Panel
    Friend WithEvents lblCommentsHead As System.Windows.Forms.Label
    Friend WithEvents lblCommentsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlCommentsDetails As System.Windows.Forms.Panel
    Friend WithEvents btnViewModifyComment As System.Windows.Forms.Button
    Friend WithEvents pnlRemediationSystems As System.Windows.Forms.Panel
    Friend WithEvents lblRemediationSystemsHead As System.Windows.Forms.Label
    Friend WithEvents lblRemediationSystemsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlRemediationSystemsDetails As System.Windows.Forms.Panel
    Friend WithEvents ugRemediationSystem As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnModifyRemediationSystem As System.Windows.Forms.Button
    Friend WithEvents pnlERACandContacts As System.Windows.Forms.Panel
    Friend WithEvents lblERACandContactsHead As System.Windows.Forms.Label
    Friend WithEvents lblERACandContactsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlERACandContactsDetails As System.Windows.Forms.Panel
    Friend WithEvents lblERAC As System.Windows.Forms.Label
    Friend WithEvents lblIRAC As System.Windows.Forms.Label
    Friend WithEvents tbPageOwnerContactList As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerContactHeader As System.Windows.Forms.Panel
    Friend WithEvents chkOwnerShowActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents lblOwnerContacts As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerContactContainer As System.Windows.Forms.Panel
    Friend WithEvents ugOwnerContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerContactButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAddSearchContact As System.Windows.Forms.Button
    Friend WithEvents tbCntrlFacility As System.Windows.Forms.TabControl
    Friend WithEvents tbPageAddLustEvent As System.Windows.Forms.TabPage
    Public WithEvents lblDateTransfered As System.Windows.Forms.Label
    Friend WithEvents lblFacilitySigNFDue As System.Windows.Forms.Label
    Public WithEvents txtDueByNF As System.Windows.Forms.TextBox
    Friend WithEvents pnlFacilityLustButton As System.Windows.Forms.Panel
    Friend WithEvents lblPMHistorytt As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerFlag As System.Windows.Forms.Button
    Friend WithEvents btnOwnerComment As System.Windows.Forms.Button
    Friend WithEvents btnFacComments As System.Windows.Forms.Button
    Friend WithEvents btnFacFlags As System.Windows.Forms.Button
    Friend WithEvents pnlERACandIRAC As System.Windows.Forms.Panel
    Friend WithEvents pnlERACContactHeader As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pnlERACContactContainer As System.Windows.Forms.Panel
    Friend WithEvents ugERACContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlERACContactButtons As System.Windows.Forms.Panel
    Friend WithEvents btnERACContactModify As System.Windows.Forms.Button
    Friend WithEvents btnERACContactDelete As System.Windows.Forms.Button
    Friend WithEvents btnERACContactAssociate As System.Windows.Forms.Button
    Friend WithEvents btnERACContactAddorSearch As System.Windows.Forms.Button
    Friend WithEvents chkERACShowActive As System.Windows.Forms.CheckBox
    Friend WithEvents chkERACShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents lblLustContacts As System.Windows.Forms.Label
    Friend WithEvents lblProject234 As System.Windows.Forms.Label
    Friend WithEvents lblProjectManager As System.Windows.Forms.Label
    Friend WithEvents chkOwnerShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkERACShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents lblERACSearch As System.Windows.Forms.Label
    Friend WithEvents lblIRACSearch As System.Windows.Forms.Label
    Friend WithEvents txtIRAC As System.Windows.Forms.TextBox
    Friend WithEvents txtERAC As System.Windows.Forms.TextBox
    Friend WithEvents btnFlagsLustEvent As System.Windows.Forms.Button
    Friend WithEvents btnGoToFinancial As System.Windows.Forms.Button
    Public WithEvents lblOwnerLastEditedOn As System.Windows.Forms.Label
    Public WithEvents lblOwnerLastEditedBy As System.Windows.Forms.Label
    Friend WithEvents lnkEnsite As System.Windows.Forms.LinkLabel
    Friend WithEvents tbPageOwnerDocuments As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityDocuments As System.Windows.Forms.TabPage
    Friend WithEvents UCOwnerDocuments As MUSTER.DocumentViewControl
    Friend WithEvents UCFacilityDocuments As MUSTER.DocumentViewControl
    Friend WithEvents pnlOwnerSummaryHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlOwnerSummaryDetails As System.Windows.Forms.Panel
    Public WithEvents UCOwnerSummary As MUSTER.OwnerSummary
    Friend WithEvents btnTecnOwnerLabels As System.Windows.Forms.Button
    Friend WithEvents btnTecOwnerEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnTecFacLabels As System.Windows.Forms.Button
    Friend WithEvents btnTecFacEnvelopes As System.Windows.Forms.Button
    Public WithEvents txtFacilitySIC As System.Windows.Forms.Label
    Friend WithEvents pnlFundsEligibilitySysQuestions As System.Windows.Forms.Panel
    Friend WithEvents lblSysQuestionsYes As System.Windows.Forms.Label
    Friend WithEvents lblSysQuestions As System.Windows.Forms.Label
    Friend WithEvents lblSysQuestionsNA As System.Windows.Forms.Label
    Friend WithEvents lblSysQuestionsNo As System.Windows.Forms.Label
    Friend WithEvents lblQuestion3 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion13NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion13No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion13Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion13 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion12NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion12No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion12Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion12 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion11NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion11No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion11Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion11 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion9NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion9No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion9Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion9 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion7NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion7No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion7Yes As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion4No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion4Yes As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion3No As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion3Yes As System.Windows.Forms.CheckBox
    Friend WithEvents lblQuestion7 As System.Windows.Forms.Label
    Friend WithEvents lblQuestion4 As System.Windows.Forms.Label
    Friend WithEvents chkQuestion4NA As System.Windows.Forms.CheckBox
    Friend WithEvents chkQuestion3NA As System.Windows.Forms.CheckBox
    Friend WithEvents lblActivitiesDocumentsCellDesc As System.Windows.Forms.Label
    Friend WithEvents cmbCause As System.Windows.Forms.ComboBox
    Friend WithEvents lblCause As System.Windows.Forms.Label
    Friend WithEvents LblAssess As System.Windows.Forms.Label
    Public WithEvents dtPickAssess As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnActPlanning As System.Windows.Forms.Button
    Friend WithEvents btnAddRemediationSystem As System.Windows.Forms.Button
    Friend WithEvents BtnSaveEngineers As System.Windows.Forms.Button
    Friend WithEvents lblCompAssDate As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Technical))
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.lblOwnerLastEditedOn = New System.Windows.Forms.Label
        Me.lblOwnerLastEditedBy = New System.Windows.Forms.Label
        Me.tbCntrlTechnical = New System.Windows.Forms.TabControl
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
        Me.tbPageOwnerDocuments = New System.Windows.Forms.TabPage
        Me.UCOwnerDocuments = New MUSTER.DocumentViewControl
        Me.pnlOwnerDetail = New System.Windows.Forms.Panel
        Me.btnTecnOwnerLabels = New System.Windows.Forms.Button
        Me.btnTecOwnerEnvelopes = New System.Windows.Forms.Button
        Me.pnlOwnerButtons = New System.Windows.Forms.Panel
        Me.btnOwnerFlag = New System.Windows.Forms.Button
        Me.btnOwnerComment = New System.Windows.Forms.Button
        Me.chkOwnerAgencyInterest = New System.Windows.Forms.CheckBox
        Me.lblOwnerActiveOrNot = New System.Windows.Forms.Label
        Me.LinkLblCAPSignup = New System.Windows.Forms.LinkLabel
        Me.lblCAPParticipationLevel = New System.Windows.Forms.Label
        Me.mskTxtOwnerFax = New AxMSMask.AxMaskEdBox
        Me.mskTxtOwnerPhone2 = New AxMSMask.AxMaskEdBox
        Me.mskTxtOwnerPhone = New AxMSMask.AxMaskEdBox
        Me.lblOwnerEmail = New System.Windows.Forms.Label
        Me.txtOwnerEmail = New System.Windows.Forms.TextBox
        Me.lblFax = New System.Windows.Forms.Label
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
        Me.tbPageFacilityDetail = New System.Windows.Forms.TabPage
        Me.pnlFacilityBottom = New System.Windows.Forms.Panel
        Me.tbCntrlFacility = New System.Windows.Forms.TabControl
        Me.tbPageAddLustEvent = New System.Windows.Forms.TabPage
        Me.btnAddLUSTEvent = New System.Windows.Forms.Button
        Me.dgLUSTEvents = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFacilityLustButton = New System.Windows.Forms.Panel
        Me.lblTotalNoOfLUSTEventsValue = New System.Windows.Forms.Label
        Me.lblTotalNoOfLUSTEvents = New System.Windows.Forms.Label
        Me.tbPageFacilityDocuments = New System.Windows.Forms.TabPage
        Me.UCFacilityDocuments = New MUSTER.DocumentViewControl
        Me.pnl_FacilityDetail = New System.Windows.Forms.Panel
        Me.dtPickAssess = New System.Windows.Forms.DateTimePicker
        Me.LblAssess = New System.Windows.Forms.Label
        Me.btnTecFacLabels = New System.Windows.Forms.Button
        Me.btnTecFacEnvelopes = New System.Windows.Forms.Button
        Me.btnFacComments = New System.Windows.Forms.Button
        Me.btnFacFlags = New System.Windows.Forms.Button
        Me.txtDueByNF = New System.Windows.Forms.TextBox
        Me.lblFacilitySigNFDue = New System.Windows.Forms.Label
        Me.lblDateTransfered = New System.Windows.Forms.Label
        Me.btnFacilityCancel = New System.Windows.Forms.Button
        Me.btnFacilitySave = New System.Windows.Forms.Button
        Me.btnTransferLUSTEvent = New System.Windows.Forms.Button
        Me.dtPickUpcomingInstallDateValue = New System.Windows.Forms.DateTimePicker
        Me.lblUpcomingInstallDate = New System.Windows.Forms.Label
        Me.chkUpcomingInstall = New System.Windows.Forms.CheckBox
        Me.lnkLblNextFac = New System.Windows.Forms.LinkLabel
        Me.lblCAPStatusValue = New System.Windows.Forms.Label
        Me.lblCAPStatus = New System.Windows.Forms.Label
        Me.txtFuelBrand = New System.Windows.Forms.TextBox
        Me.ll = New System.Windows.Forms.Label
        Me.dtFacilityPowerOff = New System.Windows.Forms.DateTimePicker
        Me.lnkLblPrevFacility = New System.Windows.Forms.LinkLabel
        Me.lblLUSTSite = New System.Windows.Forms.Label
        Me.chkLUSTSite = New System.Windows.Forms.CheckBox
        Me.lblPowerOff = New System.Windows.Forms.Label
        Me.lblCAPCandidate = New System.Windows.Forms.Label
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
        Me.lblFacilityPhone = New System.Windows.Forms.Label
        Me.txtFacilityName = New System.Windows.Forms.TextBox
        Me.lblFacilityAddress = New System.Windows.Forms.Label
        Me.lblFacilityName = New System.Windows.Forms.Label
        Me.lblFacilityLatDegree = New System.Windows.Forms.Label
        Me.txtFacilitySIC = New System.Windows.Forms.Label
        Me.tbPageLUSTEvent = New System.Windows.Forms.TabPage
        Me.pnlLustEvents = New System.Windows.Forms.Panel
        Me.pnlLUSTEventDetails = New System.Windows.Forms.Panel
        Me.pnlERACandContactsDetails = New System.Windows.Forms.Panel
        Me.pnlERACContactButtons = New System.Windows.Forms.Panel
        Me.btnERACContactModify = New System.Windows.Forms.Button
        Me.btnERACContactDelete = New System.Windows.Forms.Button
        Me.btnERACContactAssociate = New System.Windows.Forms.Button
        Me.btnERACContactAddorSearch = New System.Windows.Forms.Button
        Me.pnlERACContactContainer = New System.Windows.Forms.Panel
        Me.ugERACContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label4 = New System.Windows.Forms.Label
        Me.pnlERACContactHeader = New System.Windows.Forms.Panel
        Me.chkERACShowActive = New System.Windows.Forms.CheckBox
        Me.chkERACShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkERACShowContactsforAllModules = New System.Windows.Forms.CheckBox
        Me.lblLustContacts = New System.Windows.Forms.Label
        Me.pnlERACandIRAC = New System.Windows.Forms.Panel
        Me.BtnSaveEngineers = New System.Windows.Forms.Button
        Me.txtERAC = New System.Windows.Forms.TextBox
        Me.txtIRAC = New System.Windows.Forms.TextBox
        Me.lblIRACSearch = New System.Windows.Forms.Label
        Me.lblERACSearch = New System.Windows.Forms.Label
        Me.lblERAC = New System.Windows.Forms.Label
        Me.lblIRAC = New System.Windows.Forms.Label
        Me.pnlERACandContacts = New System.Windows.Forms.Panel
        Me.lblERACandContactsHead = New System.Windows.Forms.Label
        Me.lblERACandContactsDisplay = New System.Windows.Forms.Label
        Me.pnlRemediationSystemsDetails = New System.Windows.Forms.Panel
        Me.btnAddRemediationSystem = New System.Windows.Forms.Button
        Me.ugRemediationSystem = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnModifyRemediationSystem = New System.Windows.Forms.Button
        Me.pnlRemediationSystems = New System.Windows.Forms.Panel
        Me.lblRemediationSystemsHead = New System.Windows.Forms.Label
        Me.lblRemediationSystemsDisplay = New System.Windows.Forms.Label
        Me.pnlCommentsDetails = New System.Windows.Forms.Panel
        Me.btnViewModifyComment = New System.Windows.Forms.Button
        Me.btnFlagsLustEvent = New System.Windows.Forms.Button
        Me.pnlComments = New System.Windows.Forms.Panel
        Me.lblCommentsHead = New System.Windows.Forms.Label
        Me.lblCommentsDisplay = New System.Windows.Forms.Label
        Me.pnlActivitiesDocumentsDetails = New System.Windows.Forms.Panel
        Me.lblActivitiesDocumentsCellDesc = New System.Windows.Forms.Label
        Me.btnDeleteActivity = New System.Windows.Forms.Button
        Me.btnDeleteDocument = New System.Windows.Forms.Button
        Me.btnModifyDocument = New System.Windows.Forms.Button
        Me.btnAddDocument = New System.Windows.Forms.Button
        Me.btnModifyActivity = New System.Windows.Forms.Button
        Me.btnAddActivity = New System.Windows.Forms.Button
        Me.ugActivitiesandDocuments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.chkShowallDocuments = New System.Windows.Forms.CheckBox
        Me.chkShowallActivities = New System.Windows.Forms.CheckBox
        Me.pnlActivitiesDocuments = New System.Windows.Forms.Panel
        Me.lblActivitiesDocumentsHead = New System.Windows.Forms.Label
        Me.lblActivitiesDocumentsDisplay = New System.Windows.Forms.Label
        Me.pnlFundsEligibilityDetails = New System.Windows.Forms.Panel
        Me.pnlFundsEligibilityQuestions = New System.Windows.Forms.Panel
        Me.chkQuestion17NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion17No = New System.Windows.Forms.CheckBox
        Me.chkQuestion17Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion17 = New System.Windows.Forms.Label
        Me.chkQuestion16NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion16No = New System.Windows.Forms.CheckBox
        Me.chkQuestion16Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion16 = New System.Windows.Forms.Label
        Me.chkQuestion15NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion15No = New System.Windows.Forms.CheckBox
        Me.chkQuestion15Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion15 = New System.Windows.Forms.Label
        Me.chkQuestion14NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion14No = New System.Windows.Forms.CheckBox
        Me.chkQuestion14Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion14 = New System.Windows.Forms.Label
        Me.chkQuestion10NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion10No = New System.Windows.Forms.CheckBox
        Me.chkQuestion10Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion10 = New System.Windows.Forms.Label
        Me.chkQuestion8NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion8No = New System.Windows.Forms.CheckBox
        Me.chkQuestion8Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion8 = New System.Windows.Forms.Label
        Me.chkQuestion6NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion6No = New System.Windows.Forms.CheckBox
        Me.chkQuestion6Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion6 = New System.Windows.Forms.Label
        Me.chkQuestion5NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion5No = New System.Windows.Forms.CheckBox
        Me.chkQuestion5Yes = New System.Windows.Forms.CheckBox
        Me.chkQuestion2No = New System.Windows.Forms.CheckBox
        Me.chkQuestion2Yes = New System.Windows.Forms.CheckBox
        Me.chkQuestion1No = New System.Windows.Forms.CheckBox
        Me.chkQuestion1Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion5 = New System.Windows.Forms.Label
        Me.lblQuestion2 = New System.Windows.Forms.Label
        Me.lblQuestion1 = New System.Windows.Forms.Label
        Me.lblEligibilityComments = New System.Windows.Forms.Label
        Me.txtEligibilityComments = New System.Windows.Forms.TextBox
        Me.btnSendtoPM = New System.Windows.Forms.Button
        Me.btnViewCheckList = New System.Windows.Forms.Button
        Me.lblQuestionsNA = New System.Windows.Forms.Label
        Me.lblQuestionsNo = New System.Windows.Forms.Label
        Me.lblQuestionsYes = New System.Windows.Forms.Label
        Me.lblQuestions = New System.Windows.Forms.Label
        Me.pnlTFAssess = New System.Windows.Forms.Panel
        Me.dtCommissionOn = New System.Windows.Forms.DateTimePicker
        Me.dtOPCHeadOn = New System.Windows.Forms.DateTimePicker
        Me.dtUSTChiefOn = New System.Windows.Forms.DateTimePicker
        Me.dtPMHeadOn = New System.Windows.Forms.DateTimePicker
        Me.chkforCommission = New System.Windows.Forms.CheckBox
        Me.chkforHeadofOPC = New System.Windows.Forms.CheckBox
        Me.lblCommissionBy = New System.Windows.Forms.Label
        Me.lblOPCHeadBy = New System.Windows.Forms.Label
        Me.lblUSTChiefBy = New System.Windows.Forms.Label
        Me.txtCommissionBy = New System.Windows.Forms.TextBox
        Me.txtOPCHeadBy = New System.Windows.Forms.TextBox
        Me.txtUSTChiefBy = New System.Windows.Forms.TextBox
        Me.txtPMHeadBy = New System.Windows.Forms.TextBox
        Me.lblPMHeadBy = New System.Windows.Forms.Label
        Me.lblCommissionOn = New System.Windows.Forms.Label
        Me.lblOPCHeadOn = New System.Windows.Forms.Label
        Me.lblUSTChiefOn = New System.Windows.Forms.Label
        Me.lblPMHeadOn = New System.Windows.Forms.Label
        Me.chkCommissionNo = New System.Windows.Forms.CheckBox
        Me.chkCommissionYes = New System.Windows.Forms.CheckBox
        Me.lblCommission = New System.Windows.Forms.Label
        Me.chkOPCHeadNo = New System.Windows.Forms.CheckBox
        Me.chkOPCHeadYes = New System.Windows.Forms.CheckBox
        Me.lblOPCHead = New System.Windows.Forms.Label
        Me.chkUSTChiefUndecided = New System.Windows.Forms.CheckBox
        Me.chkUSTChiefNo = New System.Windows.Forms.CheckBox
        Me.chkUSTChiefYes = New System.Windows.Forms.CheckBox
        Me.lblUSTChief = New System.Windows.Forms.Label
        Me.lblUndecided = New System.Windows.Forms.Label
        Me.lblNo = New System.Windows.Forms.Label
        Me.lblYes = New System.Windows.Forms.Label
        Me.chkPMHeadUndecided = New System.Windows.Forms.CheckBox
        Me.chkPMHeadYes = New System.Windows.Forms.CheckBox
        Me.lblPMHead = New System.Windows.Forms.Label
        Me.pnlFundsEligibilitySysQuestions = New System.Windows.Forms.Panel
        Me.chkQuestion13NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion13No = New System.Windows.Forms.CheckBox
        Me.chkQuestion13Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion13 = New System.Windows.Forms.Label
        Me.chkQuestion12NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion12No = New System.Windows.Forms.CheckBox
        Me.chkQuestion12Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion12 = New System.Windows.Forms.Label
        Me.chkQuestion11NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion11No = New System.Windows.Forms.CheckBox
        Me.chkQuestion11Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion11 = New System.Windows.Forms.Label
        Me.chkQuestion9NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion9No = New System.Windows.Forms.CheckBox
        Me.chkQuestion9Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion9 = New System.Windows.Forms.Label
        Me.chkQuestion7NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion7No = New System.Windows.Forms.CheckBox
        Me.chkQuestion7Yes = New System.Windows.Forms.CheckBox
        Me.chkQuestion4No = New System.Windows.Forms.CheckBox
        Me.chkQuestion4Yes = New System.Windows.Forms.CheckBox
        Me.chkQuestion3No = New System.Windows.Forms.CheckBox
        Me.chkQuestion3Yes = New System.Windows.Forms.CheckBox
        Me.lblQuestion7 = New System.Windows.Forms.Label
        Me.lblQuestion4 = New System.Windows.Forms.Label
        Me.lblQuestion3 = New System.Windows.Forms.Label
        Me.chkQuestion4NA = New System.Windows.Forms.CheckBox
        Me.chkQuestion3NA = New System.Windows.Forms.CheckBox
        Me.lblSysQuestionsYes = New System.Windows.Forms.Label
        Me.lblSysQuestions = New System.Windows.Forms.Label
        Me.lblSysQuestionsNA = New System.Windows.Forms.Label
        Me.lblSysQuestionsNo = New System.Windows.Forms.Label
        Me.pnlFundsEligibility = New System.Windows.Forms.Panel
        Me.lblFundsEligibilityHead = New System.Windows.Forms.Label
        Me.lblFundsEligibilityDisplay = New System.Windows.Forms.Label
        Me.PnlReleaseInfoDetails = New System.Windows.Forms.Panel
        Me.btnEvtTankCollapse = New System.Windows.Forms.Button
        Me.btnEvtTankToggle = New System.Windows.Forms.Button
        Me.dtConfirmedOn = New System.Windows.Forms.DateTimePicker
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.chkFreeProductUnKnown = New System.Windows.Forms.CheckBox
        Me.chkFreeProductWasteOil = New System.Windows.Forms.CheckBox
        Me.chkVaporPAH = New System.Windows.Forms.CheckBox
        Me.chkVaporBTEX = New System.Windows.Forms.CheckBox
        Me.chkToCVapor = New System.Windows.Forms.CheckBox
        Me.chkFreeProductKerosene = New System.Windows.Forms.CheckBox
        Me.chkFreeProductDiesel = New System.Windows.Forms.CheckBox
        Me.chkFreeProductGasoline = New System.Windows.Forms.CheckBox
        Me.chkToCFreeProduct = New System.Windows.Forms.CheckBox
        Me.chkGroundWaterTPH = New System.Windows.Forms.CheckBox
        Me.chkGroundWaterPAH = New System.Windows.Forms.CheckBox
        Me.chkGroundWaterBTEX = New System.Windows.Forms.CheckBox
        Me.chkGroundWater = New System.Windows.Forms.CheckBox
        Me.chkSoilTPH = New System.Windows.Forms.CheckBox
        Me.chkSoilPAH = New System.Windows.Forms.CheckBox
        Me.chkSoilBTEX = New System.Windows.Forms.CheckBox
        Me.chkSoil = New System.Windows.Forms.CheckBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkInspection = New System.Windows.Forms.CheckBox
        Me.chkTankClosure = New System.Windows.Forms.CheckBox
        Me.chkInventoryShortage = New System.Windows.Forms.CheckBox
        Me.chkFailedPTT = New System.Windows.Forms.CheckBox
        Me.chkSoilContamination = New System.Windows.Forms.CheckBox
        Me.chkFreeProduct = New System.Windows.Forms.CheckBox
        Me.chkVapors = New System.Windows.Forms.CheckBox
        Me.chkGWContamination = New System.Windows.Forms.CheckBox
        Me.chkGWWell = New System.Windows.Forms.CheckBox
        Me.chkSurfaceSheen = New System.Windows.Forms.CheckBox
        Me.chkFacilityLeakDetection = New System.Windows.Forms.CheckBox
        Me.ugTankandPipes = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.cmbExtent = New System.Windows.Forms.ComboBox
        Me.lblExtent = New System.Windows.Forms.Label
        Me.cmbLocation = New System.Windows.Forms.ComboBox
        Me.cmbIdentifiedBy = New System.Windows.Forms.ComboBox
        Me.lblLocation = New System.Windows.Forms.Label
        Me.lblIdentifiedBy = New System.Windows.Forms.Label
        Me.lblConfirmedOn = New System.Windows.Forms.Label
        Me.cmbCause = New System.Windows.Forms.ComboBox
        Me.lblCause = New System.Windows.Forms.Label
        Me.pnlReleaseInfo = New System.Windows.Forms.Panel
        Me.lblReleaseInfoHead = New System.Windows.Forms.Label
        Me.lblReleaseInfoDisplay = New System.Windows.Forms.Label
        Me.PnlEventInfoDetails = New System.Windows.Forms.Panel
        Me.lblCompAssDate = New System.Windows.Forms.Label
        Me.dtCompAssDate = New System.Windows.Forms.DateTimePicker
        Me.dtLastGWS = New System.Windows.Forms.DateTimePicker
        Me.dtLastPTT = New System.Windows.Forms.DateTimePicker
        Me.dtLastLDR = New System.Windows.Forms.DateTimePicker
        Me.dtStartDate = New System.Windows.Forms.DateTimePicker
        Me.lblLastGWS = New System.Windows.Forms.Label
        Me.txtRelatedSites = New System.Windows.Forms.TextBox
        Me.lblRelatedSites = New System.Windows.Forms.Label
        Me.cmbReleaseStatus = New System.Windows.Forms.ComboBox
        Me.cmbSuspectedSource = New System.Windows.Forms.ComboBox
        Me.lblSuspectedSource = New System.Windows.Forms.Label
        Me.lblLastPTT = New System.Windows.Forms.Label
        Me.lblLastLDR = New System.Windows.Forms.Label
        Me.lblReleaseStatus = New System.Windows.Forms.Label
        Me.lblDateofReport = New System.Windows.Forms.Label
        Me.dtDateofReport = New System.Windows.Forms.DateTimePicker
        Me.lblEventStartDate = New System.Windows.Forms.Label
        Me.pnlEventInfo = New System.Windows.Forms.Panel
        Me.lblEventInfoHead = New System.Windows.Forms.Label
        Me.lblEventInfoDisplay = New System.Windows.Forms.Label
        Me.pnlLUSTEventHeader = New System.Windows.Forms.Panel
        Me.lblPMHistorytt = New System.Windows.Forms.Label
        Me.lblEventIDValue = New System.Windows.Forms.Label
        Me.lblPrioritytt = New System.Windows.Forms.Label
        Me.cmbPriority = New System.Windows.Forms.ComboBox
        Me.lblPriority = New System.Windows.Forms.Label
        Me.cmbMGPTFStatus = New System.Windows.Forms.ComboBox
        Me.lblMGPTFStatus = New System.Windows.Forms.Label
        Me.cmbEventStatus = New System.Windows.Forms.ComboBox
        Me.lblEventStatus = New System.Windows.Forms.Label
        Me.lblEventCountValue = New System.Windows.Forms.Label
        Me.lblProject234 = New System.Windows.Forms.Label
        Me.lblEventID = New System.Windows.Forms.Label
        Me.lblProjectManager = New System.Windows.Forms.Label
        Me.cmbProjectManager = New System.Windows.Forms.ComboBox
        Me.lnkEnsite = New System.Windows.Forms.LinkLabel
        Me.btnGoToFinancial = New System.Windows.Forms.Button
        Me.pnlLUSTEventBottom = New System.Windows.Forms.Panel
        Me.btnActPlanning = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnDeleteLUSTEvent = New System.Windows.Forms.Button
        Me.btnSaveLUSTEvent = New System.Windows.Forms.Button
        Me.tbPageSummary = New System.Windows.Forms.TabPage
        Me.pnlOwnerSummaryDetails = New System.Windows.Forms.Panel
        Me.UCOwnerSummary = New MUSTER.OwnerSummary
        Me.Panel12 = New System.Windows.Forms.Panel
        Me.pnlOwnerSummaryHeader = New System.Windows.Forms.Panel
        Me.pnlTop.SuspendLayout()
        Me.tbCntrlTechnical.SuspendLayout()
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
        Me.tbPageOwnerDocuments.SuspendLayout()
        Me.pnlOwnerDetail.SuspendLayout()
        Me.pnlOwnerButtons.SuspendLayout()
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageFacilityDetail.SuspendLayout()
        Me.pnlFacilityBottom.SuspendLayout()
        Me.tbCntrlFacility.SuspendLayout()
        Me.tbPageAddLustEvent.SuspendLayout()
        CType(Me.dgLUSTEvents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityLustButton.SuspendLayout()
        Me.tbPageFacilityDocuments.SuspendLayout()
        Me.pnl_FacilityDetail.SuspendLayout()
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageLUSTEvent.SuspendLayout()
        Me.pnlLustEvents.SuspendLayout()
        Me.pnlLUSTEventDetails.SuspendLayout()
        Me.pnlERACandContactsDetails.SuspendLayout()
        Me.pnlERACContactButtons.SuspendLayout()
        Me.pnlERACContactContainer.SuspendLayout()
        CType(Me.ugERACContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlERACContactHeader.SuspendLayout()
        Me.pnlERACandIRAC.SuspendLayout()
        Me.pnlERACandContacts.SuspendLayout()
        Me.pnlRemediationSystemsDetails.SuspendLayout()
        CType(Me.ugRemediationSystem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlRemediationSystems.SuspendLayout()
        Me.pnlCommentsDetails.SuspendLayout()
        Me.pnlComments.SuspendLayout()
        Me.pnlActivitiesDocumentsDetails.SuspendLayout()
        CType(Me.ugActivitiesandDocuments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlActivitiesDocuments.SuspendLayout()
        Me.pnlFundsEligibilityDetails.SuspendLayout()
        Me.pnlFundsEligibilityQuestions.SuspendLayout()
        Me.pnlTFAssess.SuspendLayout()
        Me.pnlFundsEligibilitySysQuestions.SuspendLayout()
        Me.pnlFundsEligibility.SuspendLayout()
        Me.PnlReleaseInfoDetails.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.ugTankandPipes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlReleaseInfo.SuspendLayout()
        Me.PnlEventInfoDetails.SuspendLayout()
        Me.pnlEventInfo.SuspendLayout()
        Me.pnlLUSTEventHeader.SuspendLayout()
        Me.pnlLUSTEventBottom.SuspendLayout()
        Me.tbPageSummary.SuspendLayout()
        Me.pnlOwnerSummaryDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.BackColor = System.Drawing.SystemColors.Control
        Me.pnlTop.Controls.Add(Me.lblOwnerLastEditedOn)
        Me.pnlTop.Controls.Add(Me.lblOwnerLastEditedBy)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(1024, 24)
        Me.pnlTop.TabIndex = 1
        '
        'lblOwnerLastEditedOn
        '
        Me.lblOwnerLastEditedOn.Location = New System.Drawing.Point(744, 5)
        Me.lblOwnerLastEditedOn.Name = "lblOwnerLastEditedOn"
        Me.lblOwnerLastEditedOn.Size = New System.Drawing.Size(168, 16)
        Me.lblOwnerLastEditedOn.TabIndex = 1020
        Me.lblOwnerLastEditedOn.Text = "Last Edited On :"
        '
        'lblOwnerLastEditedBy
        '
        Me.lblOwnerLastEditedBy.Location = New System.Drawing.Point(528, 4)
        Me.lblOwnerLastEditedBy.Name = "lblOwnerLastEditedBy"
        Me.lblOwnerLastEditedBy.Size = New System.Drawing.Size(208, 16)
        Me.lblOwnerLastEditedBy.TabIndex = 1019
        Me.lblOwnerLastEditedBy.Text = "Last Edited By :"
        '
        'tbCntrlTechnical
        '
        Me.tbCntrlTechnical.Controls.Add(Me.tbPageOwnerDetail)
        Me.tbCntrlTechnical.Controls.Add(Me.tbPageFacilityDetail)
        Me.tbCntrlTechnical.Controls.Add(Me.tbPageLUSTEvent)
        Me.tbCntrlTechnical.Controls.Add(Me.tbPageSummary)
        Me.tbCntrlTechnical.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCntrlTechnical.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCntrlTechnical.ItemSize = New System.Drawing.Size(64, 18)
        Me.tbCntrlTechnical.Location = New System.Drawing.Point(0, 24)
        Me.tbCntrlTechnical.Multiline = True
        Me.tbCntrlTechnical.Name = "tbCntrlTechnical"
        Me.tbCntrlTechnical.SelectedIndex = 0
        Me.tbCntrlTechnical.ShowToolTips = True
        Me.tbCntrlTechnical.Size = New System.Drawing.Size(1024, 670)
        Me.tbCntrlTechnical.TabIndex = 2
        '
        'tbPageOwnerDetail
        '
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerBottom)
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerDetail)
        Me.tbPageOwnerDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageOwnerDetail.Name = "tbPageOwnerDetail"
        Me.tbPageOwnerDetail.Size = New System.Drawing.Size(1016, 644)
        Me.tbPageOwnerDetail.TabIndex = 7
        Me.tbPageOwnerDetail.Text = "Owner Details"
        '
        'pnlOwnerBottom
        '
        Me.pnlOwnerBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlOwnerBottom.Controls.Add(Me.tbCtrlOwner)
        Me.pnlOwnerBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerBottom.Location = New System.Drawing.Point(0, 200)
        Me.pnlOwnerBottom.Name = "pnlOwnerBottom"
        Me.pnlOwnerBottom.Size = New System.Drawing.Size(1016, 444)
        Me.pnlOwnerBottom.TabIndex = 44
        '
        'tbCtrlOwner
        '
        Me.tbCtrlOwner.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerFacilities)
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerContactList)
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerDocuments)
        Me.tbCtrlOwner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlOwner.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlOwner.Name = "tbCtrlOwner"
        Me.tbCtrlOwner.SelectedIndex = 0
        Me.tbCtrlOwner.Size = New System.Drawing.Size(1014, 442)
        Me.tbCtrlOwner.TabIndex = 9
        '
        'tbPageOwnerFacilities
        '
        Me.tbPageOwnerFacilities.BackColor = System.Drawing.SystemColors.Control
        Me.tbPageOwnerFacilities.Controls.Add(Me.ugFacilityList)
        Me.tbPageOwnerFacilities.Controls.Add(Me.pnlOwnerFacilityBottom)
        Me.tbPageOwnerFacilities.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerFacilities.Name = "tbPageOwnerFacilities"
        Me.tbPageOwnerFacilities.Size = New System.Drawing.Size(1006, 411)
        Me.tbPageOwnerFacilities.TabIndex = 0
        Me.tbPageOwnerFacilities.Text = "Facilities"
        '
        'ugFacilityList
        '
        Me.ugFacilityList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilityList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugFacilityList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFacilityList.Location = New System.Drawing.Point(0, 0)
        Me.ugFacilityList.Name = "ugFacilityList"
        Me.ugFacilityList.Size = New System.Drawing.Size(1006, 387)
        Me.ugFacilityList.TabIndex = 20
        '
        'pnlOwnerFacilityBottom
        '
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.lblNoOfFacilitiesValue)
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.lblNoOfFacilities)
        Me.pnlOwnerFacilityBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerFacilityBottom.Location = New System.Drawing.Point(0, 387)
        Me.pnlOwnerFacilityBottom.Name = "pnlOwnerFacilityBottom"
        Me.pnlOwnerFacilityBottom.Size = New System.Drawing.Size(1006, 24)
        Me.pnlOwnerFacilityBottom.TabIndex = 19
        '
        'lblNoOfFacilitiesValue
        '
        Me.lblNoOfFacilitiesValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilitiesValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfFacilitiesValue.Location = New System.Drawing.Point(100, 0)
        Me.lblNoOfFacilitiesValue.Name = "lblNoOfFacilitiesValue"
        Me.lblNoOfFacilitiesValue.Size = New System.Drawing.Size(56, 24)
        Me.lblNoOfFacilitiesValue.TabIndex = 7
        Me.lblNoOfFacilitiesValue.Text = "0"
        '
        'lblNoOfFacilities
        '
        Me.lblNoOfFacilities.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilities.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfFacilities.Location = New System.Drawing.Point(0, 0)
        Me.lblNoOfFacilities.Name = "lblNoOfFacilities"
        Me.lblNoOfFacilities.Size = New System.Drawing.Size(100, 24)
        Me.lblNoOfFacilities.TabIndex = 6
        Me.lblNoOfFacilities.Text = "No of Facilities:"
        '
        'tbPageOwnerContactList
        '
        Me.tbPageOwnerContactList.AutoScroll = True
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactContainer)
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactHeader)
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactButtons)
        Me.tbPageOwnerContactList.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerContactList.Name = "tbPageOwnerContactList"
        Me.tbPageOwnerContactList.Size = New System.Drawing.Size(1006, 411)
        Me.tbPageOwnerContactList.TabIndex = 1
        Me.tbPageOwnerContactList.Text = "Contacts"
        '
        'pnlOwnerContactContainer
        '
        Me.pnlOwnerContactContainer.Controls.Add(Me.ugOwnerContacts)
        Me.pnlOwnerContactContainer.Controls.Add(Me.Label1)
        Me.pnlOwnerContactContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerContactContainer.Location = New System.Drawing.Point(0, 25)
        Me.pnlOwnerContactContainer.Name = "pnlOwnerContactContainer"
        Me.pnlOwnerContactContainer.Size = New System.Drawing.Size(1006, 356)
        Me.pnlOwnerContactContainer.TabIndex = 2
        '
        'ugOwnerContacts
        '
        Me.ugOwnerContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugOwnerContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugOwnerContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugOwnerContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugOwnerContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugOwnerContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugOwnerContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugOwnerContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugOwnerContacts.Name = "ugOwnerContacts"
        Me.ugOwnerContacts.Size = New System.Drawing.Size(1006, 356)
        Me.ugOwnerContacts.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(792, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(7, 23)
        Me.Label1.TabIndex = 2
        '
        'pnlOwnerContactHeader
        '
        Me.pnlOwnerContactHeader.Controls.Add(Me.chkOwnerShowActiveOnly)
        Me.pnlOwnerContactHeader.Controls.Add(Me.chkOwnerShowRelatedContacts)
        Me.pnlOwnerContactHeader.Controls.Add(Me.chkOwnerShowContactsforAllModules)
        Me.pnlOwnerContactHeader.Controls.Add(Me.lblOwnerContacts)
        Me.pnlOwnerContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerContactHeader.DockPadding.All = 3
        Me.pnlOwnerContactHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerContactHeader.Name = "pnlOwnerContactHeader"
        Me.pnlOwnerContactHeader.Size = New System.Drawing.Size(1006, 25)
        Me.pnlOwnerContactHeader.TabIndex = 0
        '
        'chkOwnerShowActiveOnly
        '
        Me.chkOwnerShowActiveOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowActiveOnly.Location = New System.Drawing.Point(635, 6)
        Me.chkOwnerShowActiveOnly.Name = "chkOwnerShowActiveOnly"
        Me.chkOwnerShowActiveOnly.Size = New System.Drawing.Size(144, 16)
        Me.chkOwnerShowActiveOnly.TabIndex = 2
        Me.chkOwnerShowActiveOnly.Tag = "646"
        Me.chkOwnerShowActiveOnly.Text = "Show Active Only"
        '
        'chkOwnerShowRelatedContacts
        '
        Me.chkOwnerShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowRelatedContacts.Location = New System.Drawing.Point(467, 6)
        Me.chkOwnerShowRelatedContacts.Name = "chkOwnerShowRelatedContacts"
        Me.chkOwnerShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkOwnerShowRelatedContacts.TabIndex = 1
        Me.chkOwnerShowRelatedContacts.Tag = "645"
        Me.chkOwnerShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkOwnerShowContactsforAllModules
        '
        Me.chkOwnerShowContactsforAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowContactsforAllModules.Location = New System.Drawing.Point(251, 6)
        Me.chkOwnerShowContactsforAllModules.Name = "chkOwnerShowContactsforAllModules"
        Me.chkOwnerShowContactsforAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkOwnerShowContactsforAllModules.TabIndex = 0
        Me.chkOwnerShowContactsforAllModules.Tag = "644"
        Me.chkOwnerShowContactsforAllModules.Text = "Show Contacts for All Modules"
        '
        'lblOwnerContacts
        '
        Me.lblOwnerContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerContacts.Location = New System.Drawing.Point(8, 6)
        Me.lblOwnerContacts.Name = "lblOwnerContacts"
        Me.lblOwnerContacts.Size = New System.Drawing.Size(104, 16)
        Me.lblOwnerContacts.TabIndex = 139
        Me.lblOwnerContacts.Text = "Owner Contacts"
        '
        'pnlOwnerContactButtons
        '
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerModifyContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerDeleteContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerAssociateContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerAddSearchContact)
        Me.pnlOwnerContactButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerContactButtons.DockPadding.All = 3
        Me.pnlOwnerContactButtons.Location = New System.Drawing.Point(0, 381)
        Me.pnlOwnerContactButtons.Name = "pnlOwnerContactButtons"
        Me.pnlOwnerContactButtons.Size = New System.Drawing.Size(1006, 30)
        Me.pnlOwnerContactButtons.TabIndex = 3
        '
        'btnOwnerModifyContact
        '
        Me.btnOwnerModifyContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerModifyContact.Location = New System.Drawing.Point(240, 4)
        Me.btnOwnerModifyContact.Name = "btnOwnerModifyContact"
        Me.btnOwnerModifyContact.Size = New System.Drawing.Size(235, 26)
        Me.btnOwnerModifyContact.TabIndex = 1
        Me.btnOwnerModifyContact.Text = "Modify Contact"
        '
        'btnOwnerDeleteContact
        '
        Me.btnOwnerDeleteContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerDeleteContact.Location = New System.Drawing.Point(472, 4)
        Me.btnOwnerDeleteContact.Name = "btnOwnerDeleteContact"
        Me.btnOwnerDeleteContact.Size = New System.Drawing.Size(235, 26)
        Me.btnOwnerDeleteContact.TabIndex = 2
        Me.btnOwnerDeleteContact.Text = "Disassociate Contact"
        '
        'btnOwnerAssociateContact
        '
        Me.btnOwnerAssociateContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerAssociateContact.Location = New System.Drawing.Point(704, 4)
        Me.btnOwnerAssociateContact.Name = "btnOwnerAssociateContact"
        Me.btnOwnerAssociateContact.Size = New System.Drawing.Size(235, 26)
        Me.btnOwnerAssociateContact.TabIndex = 3
        Me.btnOwnerAssociateContact.Text = "Associate Contact from Different Module"
        '
        'btnOwnerAddSearchContact
        '
        Me.btnOwnerAddSearchContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerAddSearchContact.Location = New System.Drawing.Point(8, 4)
        Me.btnOwnerAddSearchContact.Name = "btnOwnerAddSearchContact"
        Me.btnOwnerAddSearchContact.Size = New System.Drawing.Size(235, 26)
        Me.btnOwnerAddSearchContact.TabIndex = 0
        Me.btnOwnerAddSearchContact.Text = "Add/Search Contact to Associate"
        '
        'tbPageOwnerDocuments
        '
        Me.tbPageOwnerDocuments.Controls.Add(Me.UCOwnerDocuments)
        Me.tbPageOwnerDocuments.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerDocuments.Name = "tbPageOwnerDocuments"
        Me.tbPageOwnerDocuments.Size = New System.Drawing.Size(1006, 411)
        Me.tbPageOwnerDocuments.TabIndex = 2
        Me.tbPageOwnerDocuments.Text = "Documents"
        '
        'UCOwnerDocuments
        '
        Me.UCOwnerDocuments.AutoScroll = True
        Me.UCOwnerDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerDocuments.Name = "UCOwnerDocuments"
        Me.UCOwnerDocuments.Size = New System.Drawing.Size(1006, 411)
        Me.UCOwnerDocuments.TabIndex = 3
        '
        'pnlOwnerDetail
        '
        Me.pnlOwnerDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnlOwnerDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlOwnerDetail.Controls.Add(Me.btnTecnOwnerLabels)
        Me.pnlOwnerDetail.Controls.Add(Me.btnTecOwnerEnvelopes)
        Me.pnlOwnerDetail.Controls.Add(Me.pnlOwnerButtons)
        Me.pnlOwnerDetail.Controls.Add(Me.chkOwnerAgencyInterest)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerActiveOrNot)
        Me.pnlOwnerDetail.Controls.Add(Me.LinkLblCAPSignup)
        Me.pnlOwnerDetail.Controls.Add(Me.lblCAPParticipationLevel)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerFax)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerPhone2)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerPhone)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerEmail)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerEmail)
        Me.pnlOwnerDetail.Controls.Add(Me.lblFax)
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
        Me.pnlOwnerDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerDetail.Name = "pnlOwnerDetail"
        Me.pnlOwnerDetail.Size = New System.Drawing.Size(1016, 200)
        Me.pnlOwnerDetail.TabIndex = 0
        '
        'btnTecnOwnerLabels
        '
        Me.btnTecnOwnerLabels.Location = New System.Drawing.Point(8, 112)
        Me.btnTecnOwnerLabels.Name = "btnTecnOwnerLabels"
        Me.btnTecnOwnerLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnTecnOwnerLabels.TabIndex = 1066
        Me.btnTecnOwnerLabels.Text = "Labels"
        '
        'btnTecOwnerEnvelopes
        '
        Me.btnTecOwnerEnvelopes.Location = New System.Drawing.Point(8, 84)
        Me.btnTecOwnerEnvelopes.Name = "btnTecOwnerEnvelopes"
        Me.btnTecOwnerEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnTecOwnerEnvelopes.TabIndex = 1065
        Me.btnTecOwnerEnvelopes.Text = "Envelopes"
        '
        'pnlOwnerButtons
        '
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerFlag)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerComment)
        Me.pnlOwnerButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.pnlOwnerButtons.Location = New System.Drawing.Point(336, 144)
        Me.pnlOwnerButtons.Name = "pnlOwnerButtons"
        Me.pnlOwnerButtons.Size = New System.Drawing.Size(192, 36)
        Me.pnlOwnerButtons.TabIndex = 1007
        '
        'btnOwnerFlag
        '
        Me.btnOwnerFlag.Location = New System.Drawing.Point(8, 7)
        Me.btnOwnerFlag.Name = "btnOwnerFlag"
        Me.btnOwnerFlag.Size = New System.Drawing.Size(75, 26)
        Me.btnOwnerFlag.TabIndex = 48
        Me.btnOwnerFlag.Text = "Flags"
        '
        'btnOwnerComment
        '
        Me.btnOwnerComment.Location = New System.Drawing.Point(88, 7)
        Me.btnOwnerComment.Name = "btnOwnerComment"
        Me.btnOwnerComment.Size = New System.Drawing.Size(80, 26)
        Me.btnOwnerComment.TabIndex = 47
        Me.btnOwnerComment.Text = "Comments"
        '
        'chkOwnerAgencyInterest
        '
        Me.chkOwnerAgencyInterest.Enabled = False
        Me.chkOwnerAgencyInterest.Location = New System.Drawing.Point(560, 24)
        Me.chkOwnerAgencyInterest.Name = "chkOwnerAgencyInterest"
        Me.chkOwnerAgencyInterest.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkOwnerAgencyInterest.Size = New System.Drawing.Size(112, 20)
        Me.chkOwnerAgencyInterest.TabIndex = 7
        Me.chkOwnerAgencyInterest.Text = "Agency Interest   "
        Me.chkOwnerAgencyInterest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOwnerActiveOrNot
        '
        Me.lblOwnerActiveOrNot.BackColor = System.Drawing.SystemColors.Control
        Me.lblOwnerActiveOrNot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblOwnerActiveOrNot.Enabled = False
        Me.lblOwnerActiveOrNot.Location = New System.Drawing.Point(432, 8)
        Me.lblOwnerActiveOrNot.Name = "lblOwnerActiveOrNot"
        Me.lblOwnerActiveOrNot.Size = New System.Drawing.Size(96, 20)
        Me.lblOwnerActiveOrNot.TabIndex = 1006
        '
        'LinkLblCAPSignup
        '
        Me.LinkLblCAPSignup.Enabled = False
        Me.LinkLblCAPSignup.Location = New System.Drawing.Point(560, 88)
        Me.LinkLblCAPSignup.Name = "LinkLblCAPSignup"
        Me.LinkLblCAPSignup.Size = New System.Drawing.Size(152, 16)
        Me.LinkLblCAPSignup.TabIndex = 1005
        Me.LinkLblCAPSignup.TabStop = True
        Me.LinkLblCAPSignup.Text = "CAP Signup/Maintenance"
        '
        'lblCAPParticipationLevel
        '
        Me.lblCAPParticipationLevel.Location = New System.Drawing.Point(688, 56)
        Me.lblCAPParticipationLevel.Name = "lblCAPParticipationLevel"
        Me.lblCAPParticipationLevel.Size = New System.Drawing.Size(264, 20)
        Me.lblCAPParticipationLevel.TabIndex = 1004
        Me.lblCAPParticipationLevel.Text = "NONE - 0/0 (Compliant/Candidate)"
        Me.lblCAPParticipationLevel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'mskTxtOwnerFax
        '
        Me.mskTxtOwnerFax.ContainingControl = Me
        Me.mskTxtOwnerFax.Location = New System.Drawing.Point(432, 104)
        Me.mskTxtOwnerFax.Name = "mskTxtOwnerFax"
        Me.mskTxtOwnerFax.OcxState = CType(resources.GetObject("mskTxtOwnerFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerFax.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerFax.TabIndex = 6
        '
        'mskTxtOwnerPhone2
        '
        Me.mskTxtOwnerPhone2.ContainingControl = Me
        Me.mskTxtOwnerPhone2.Location = New System.Drawing.Point(432, 80)
        Me.mskTxtOwnerPhone2.Name = "mskTxtOwnerPhone2"
        Me.mskTxtOwnerPhone2.OcxState = CType(resources.GetObject("mskTxtOwnerPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone2.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerPhone2.TabIndex = 5
        '
        'mskTxtOwnerPhone
        '
        Me.mskTxtOwnerPhone.ContainingControl = Me
        Me.mskTxtOwnerPhone.Location = New System.Drawing.Point(432, 56)
        Me.mskTxtOwnerPhone.Name = "mskTxtOwnerPhone"
        Me.mskTxtOwnerPhone.OcxState = CType(resources.GetObject("mskTxtOwnerPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerPhone.TabIndex = 4
        '
        'lblOwnerEmail
        '
        Me.lblOwnerEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerEmail.Location = New System.Drawing.Point(560, 120)
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
        Me.txtOwnerEmail.Enabled = False
        Me.txtOwnerEmail.Location = New System.Drawing.Point(608, 120)
        Me.txtOwnerEmail.Name = "txtOwnerEmail"
        Me.txtOwnerEmail.Size = New System.Drawing.Size(200, 20)
        Me.txtOwnerEmail.TabIndex = 8
        Me.txtOwnerEmail.Text = ""
        Me.txtOwnerEmail.WordWrap = False
        '
        'lblFax
        '
        Me.lblFax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFax.Location = New System.Drawing.Point(346, 104)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(80, 20)
        Me.lblFax.TabIndex = 44
        Me.lblFax.Text = "Fax:"
        Me.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOwnerAddress
        '
        Me.txtOwnerAddress.Location = New System.Drawing.Point(80, 56)
        Me.txtOwnerAddress.Multiline = True
        Me.txtOwnerAddress.Name = "txtOwnerAddress"
        Me.txtOwnerAddress.ReadOnly = True
        Me.txtOwnerAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtOwnerAddress.Size = New System.Drawing.Size(248, 103)
        Me.txtOwnerAddress.TabIndex = 1
        Me.txtOwnerAddress.Text = ""
        Me.txtOwnerAddress.WordWrap = False
        '
        'lblOwnerAddress
        '
        Me.lblOwnerAddress.Location = New System.Drawing.Point(8, 56)
        Me.lblOwnerAddress.Name = "lblOwnerAddress"
        Me.lblOwnerAddress.Size = New System.Drawing.Size(64, 20)
        Me.lblOwnerAddress.TabIndex = 88
        Me.lblOwnerAddress.Text = "Address:"
        Me.lblOwnerAddress.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOwnerName
        '
        Me.txtOwnerName.Location = New System.Drawing.Point(80, 32)
        Me.txtOwnerName.Name = "txtOwnerName"
        Me.txtOwnerName.ReadOnly = True
        Me.txtOwnerName.Size = New System.Drawing.Size(248, 21)
        Me.txtOwnerName.TabIndex = 0
        Me.txtOwnerName.Text = ""
        '
        'lblOwnerName
        '
        Me.lblOwnerName.Location = New System.Drawing.Point(8, 32)
        Me.lblOwnerName.Name = "lblOwnerName"
        Me.lblOwnerName.Size = New System.Drawing.Size(64, 20)
        Me.lblOwnerName.TabIndex = 86
        Me.lblOwnerName.Text = "Name:"
        Me.lblOwnerName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOwnerStatus
        '
        Me.lblOwnerStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerStatus.Location = New System.Drawing.Point(344, 8)
        Me.lblOwnerStatus.Name = "lblOwnerStatus"
        Me.lblOwnerStatus.Size = New System.Drawing.Size(80, 20)
        Me.lblOwnerStatus.TabIndex = 84
        Me.lblOwnerStatus.Text = "Owner Status"
        '
        'lblOwnerCapParticipant
        '
        Me.lblOwnerCapParticipant.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerCapParticipant.Location = New System.Drawing.Point(560, 56)
        Me.lblOwnerCapParticipant.Name = "lblOwnerCapParticipant"
        Me.lblOwnerCapParticipant.Size = New System.Drawing.Size(128, 20)
        Me.lblOwnerCapParticipant.TabIndex = 52
        Me.lblOwnerCapParticipant.Text = "CAP Participation Level"
        Me.lblOwnerCapParticipant.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPhone2
        '
        Me.lblPhone2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhone2.Location = New System.Drawing.Point(344, 80)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(80, 20)
        Me.lblPhone2.TabIndex = 45
        Me.lblPhone2.Text = "Phone 2:"
        Me.lblPhone2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOwnerType
        '
        Me.lblOwnerType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerType.Location = New System.Drawing.Point(8, 160)
        Me.lblOwnerType.Name = "lblOwnerType"
        Me.lblOwnerType.Size = New System.Drawing.Size(72, 20)
        Me.lblOwnerType.TabIndex = 40
        Me.lblOwnerType.Text = "Owner Type:"
        Me.lblOwnerType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOwnerAIID
        '
        Me.txtOwnerAIID.AcceptsTab = True
        Me.txtOwnerAIID.AutoSize = False
        Me.txtOwnerAIID.Enabled = False
        Me.txtOwnerAIID.Location = New System.Drawing.Point(432, 32)
        Me.txtOwnerAIID.Name = "txtOwnerAIID"
        Me.txtOwnerAIID.Size = New System.Drawing.Size(96, 21)
        Me.txtOwnerAIID.TabIndex = 3
        Me.txtOwnerAIID.Text = ""
        Me.txtOwnerAIID.WordWrap = False
        '
        'lblOwnerAIID
        '
        Me.lblOwnerAIID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerAIID.Location = New System.Drawing.Point(344, 32)
        Me.lblOwnerAIID.Name = "lblOwnerAIID"
        Me.lblOwnerAIID.Size = New System.Drawing.Size(80, 20)
        Me.lblOwnerAIID.TabIndex = 38
        Me.lblOwnerAIID.Text = "Ensite ID:"
        Me.lblOwnerAIID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOwnerIDValue
        '
        Me.lblOwnerIDValue.Enabled = False
        Me.lblOwnerIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerIDValue.Location = New System.Drawing.Point(86, 8)
        Me.lblOwnerIDValue.Name = "lblOwnerIDValue"
        Me.lblOwnerIDValue.Size = New System.Drawing.Size(96, 20)
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
        Me.lblOwnerID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOwnerPhone
        '
        Me.lblOwnerPhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerPhone.Location = New System.Drawing.Point(344, 56)
        Me.lblOwnerPhone.Name = "lblOwnerPhone"
        Me.lblOwnerPhone.Size = New System.Drawing.Size(80, 20)
        Me.lblOwnerPhone.TabIndex = 32
        Me.lblOwnerPhone.Text = "Phone:"
        Me.lblOwnerPhone.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbOwnerType
        '
        Me.cmbOwnerType.DisplayMember = "1"
        Me.cmbOwnerType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerType.DropDownWidth = 200
        Me.cmbOwnerType.Enabled = False
        Me.cmbOwnerType.ItemHeight = 15
        Me.cmbOwnerType.Location = New System.Drawing.Point(80, 160)
        Me.cmbOwnerType.Name = "cmbOwnerType"
        Me.cmbOwnerType.Size = New System.Drawing.Size(248, 23)
        Me.cmbOwnerType.TabIndex = 2
        Me.cmbOwnerType.ValueMember = "1"
        '
        'tbPageFacilityDetail
        '
        Me.tbPageFacilityDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageFacilityDetail.Controls.Add(Me.pnlFacilityBottom)
        Me.tbPageFacilityDetail.Controls.Add(Me.pnl_FacilityDetail)
        Me.tbPageFacilityDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageFacilityDetail.Name = "tbPageFacilityDetail"
        Me.tbPageFacilityDetail.Size = New System.Drawing.Size(1016, 644)
        Me.tbPageFacilityDetail.TabIndex = 8
        Me.tbPageFacilityDetail.Text = "Facility Details"
        Me.tbPageFacilityDetail.Visible = False
        '
        'pnlFacilityBottom
        '
        Me.pnlFacilityBottom.AutoScroll = True
        Me.pnlFacilityBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlFacilityBottom.Controls.Add(Me.tbCntrlFacility)
        Me.pnlFacilityBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityBottom.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlFacilityBottom.Location = New System.Drawing.Point(0, 272)
        Me.pnlFacilityBottom.Name = "pnlFacilityBottom"
        Me.pnlFacilityBottom.Size = New System.Drawing.Size(1012, 368)
        Me.pnlFacilityBottom.TabIndex = 3
        '
        'tbCntrlFacility
        '
        Me.tbCntrlFacility.Controls.Add(Me.tbPageAddLustEvent)
        Me.tbCntrlFacility.Controls.Add(Me.tbPageFacilityDocuments)
        Me.tbCntrlFacility.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCntrlFacility.Location = New System.Drawing.Point(0, 0)
        Me.tbCntrlFacility.Name = "tbCntrlFacility"
        Me.tbCntrlFacility.SelectedIndex = 0
        Me.tbCntrlFacility.Size = New System.Drawing.Size(1010, 366)
        Me.tbCntrlFacility.TabIndex = 31
        '
        'tbPageAddLustEvent
        '
        Me.tbPageAddLustEvent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageAddLustEvent.Controls.Add(Me.btnAddLUSTEvent)
        Me.tbPageAddLustEvent.Controls.Add(Me.dgLUSTEvents)
        Me.tbPageAddLustEvent.Controls.Add(Me.pnlFacilityLustButton)
        Me.tbPageAddLustEvent.Location = New System.Drawing.Point(4, 24)
        Me.tbPageAddLustEvent.Name = "tbPageAddLustEvent"
        Me.tbPageAddLustEvent.Size = New System.Drawing.Size(1002, 338)
        Me.tbPageAddLustEvent.TabIndex = 0
        Me.tbPageAddLustEvent.Text = "Lust Event"
        '
        'btnAddLUSTEvent
        '
        Me.btnAddLUSTEvent.Location = New System.Drawing.Point(0, 0)
        Me.btnAddLUSTEvent.Name = "btnAddLUSTEvent"
        Me.btnAddLUSTEvent.Size = New System.Drawing.Size(120, 24)
        Me.btnAddLUSTEvent.TabIndex = 27
        Me.btnAddLUSTEvent.Text = "Add LUST Event"
        Me.btnAddLUSTEvent.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dgLUSTEvents
        '
        Me.dgLUSTEvents.Cursor = System.Windows.Forms.Cursors.Default
        Me.dgLUSTEvents.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.dgLUSTEvents.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.dgLUSTEvents.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.dgLUSTEvents.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.dgLUSTEvents.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.dgLUSTEvents.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgLUSTEvents.Location = New System.Drawing.Point(0, 0)
        Me.dgLUSTEvents.Name = "dgLUSTEvents"
        Me.dgLUSTEvents.Size = New System.Drawing.Size(998, 310)
        Me.dgLUSTEvents.TabIndex = 30
        Me.dgLUSTEvents.Text = "Lust Events"
        '
        'pnlFacilityLustButton
        '
        Me.pnlFacilityLustButton.Controls.Add(Me.lblTotalNoOfLUSTEventsValue)
        Me.pnlFacilityLustButton.Controls.Add(Me.lblTotalNoOfLUSTEvents)
        Me.pnlFacilityLustButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFacilityLustButton.Location = New System.Drawing.Point(0, 310)
        Me.pnlFacilityLustButton.Name = "pnlFacilityLustButton"
        Me.pnlFacilityLustButton.Size = New System.Drawing.Size(998, 24)
        Me.pnlFacilityLustButton.TabIndex = 97
        '
        'lblTotalNoOfLUSTEventsValue
        '
        Me.lblTotalNoOfLUSTEventsValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfLUSTEventsValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfLUSTEventsValue.Location = New System.Drawing.Point(144, 0)
        Me.lblTotalNoOfLUSTEventsValue.Name = "lblTotalNoOfLUSTEventsValue"
        Me.lblTotalNoOfLUSTEventsValue.Size = New System.Drawing.Size(48, 24)
        Me.lblTotalNoOfLUSTEventsValue.TabIndex = 5
        Me.lblTotalNoOfLUSTEventsValue.Text = "0"
        '
        'lblTotalNoOfLUSTEvents
        '
        Me.lblTotalNoOfLUSTEvents.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfLUSTEvents.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfLUSTEvents.Location = New System.Drawing.Point(0, 0)
        Me.lblTotalNoOfLUSTEvents.Name = "lblTotalNoOfLUSTEvents"
        Me.lblTotalNoOfLUSTEvents.Size = New System.Drawing.Size(144, 24)
        Me.lblTotalNoOfLUSTEvents.TabIndex = 4
        Me.lblTotalNoOfLUSTEvents.Text = "Number of LUST Events:"
        '
        'tbPageFacilityDocuments
        '
        Me.tbPageFacilityDocuments.Controls.Add(Me.UCFacilityDocuments)
        Me.tbPageFacilityDocuments.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilityDocuments.Name = "tbPageFacilityDocuments"
        Me.tbPageFacilityDocuments.Size = New System.Drawing.Size(1002, 338)
        Me.tbPageFacilityDocuments.TabIndex = 1
        Me.tbPageFacilityDocuments.Text = "Documents"
        '
        'UCFacilityDocuments
        '
        Me.UCFacilityDocuments.AutoScroll = True
        Me.UCFacilityDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCFacilityDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCFacilityDocuments.Name = "UCFacilityDocuments"
        Me.UCFacilityDocuments.Size = New System.Drawing.Size(1002, 338)
        Me.UCFacilityDocuments.TabIndex = 3
        '
        'pnl_FacilityDetail
        '
        Me.pnl_FacilityDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnl_FacilityDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickAssess)
        Me.pnl_FacilityDetail.Controls.Add(Me.LblAssess)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnTecFacLabels)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnTecFacEnvelopes)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacComments)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacFlags)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtDueByNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilitySigNFDue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateTransfered)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacilityCancel)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacilitySave)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnTransferLUSTEvent)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickUpcomingInstallDateValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblUpcomingInstallDate)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkUpcomingInstall)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblNextFac)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPStatusValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPStatus)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFuelBrand)
        Me.pnl_FacilityDetail.Controls.Add(Me.ll)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtFacilityPowerOff)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblPrevFacility)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblPowerOff)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPCandidate)
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
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilitySIC)
        Me.pnl_FacilityDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_FacilityDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnl_FacilityDetail.Name = "pnl_FacilityDetail"
        Me.pnl_FacilityDetail.Size = New System.Drawing.Size(1012, 272)
        Me.pnl_FacilityDetail.TabIndex = 1
        '
        'dtPickAssess
        '
        Me.dtPickAssess.Checked = False
        Me.dtPickAssess.Enabled = False
        Me.dtPickAssess.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickAssess.Location = New System.Drawing.Point(816, 176)
        Me.dtPickAssess.Name = "dtPickAssess"
        Me.dtPickAssess.ShowCheckBox = True
        Me.dtPickAssess.Size = New System.Drawing.Size(112, 21)
        Me.dtPickAssess.TabIndex = 1067
        '
        'LblAssess
        '
        Me.LblAssess.Location = New System.Drawing.Point(664, 176)
        Me.LblAssess.Name = "LblAssess"
        Me.LblAssess.Size = New System.Drawing.Size(144, 20)
        Me.LblAssess.TabIndex = 1068
        Me.LblAssess.Text = "TOS Assessment Date:"
        Me.LblAssess.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnTecFacLabels
        '
        Me.btnTecFacLabels.Location = New System.Drawing.Point(8, 112)
        Me.btnTecFacLabels.Name = "btnTecFacLabels"
        Me.btnTecFacLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnTecFacLabels.TabIndex = 1066
        Me.btnTecFacLabels.Text = "Labels"
        '
        'btnTecFacEnvelopes
        '
        Me.btnTecFacEnvelopes.Location = New System.Drawing.Point(8, 84)
        Me.btnTecFacEnvelopes.Name = "btnTecFacEnvelopes"
        Me.btnTecFacEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnTecFacEnvelopes.TabIndex = 1065
        Me.btnTecFacEnvelopes.Text = "Envelopes"
        '
        'btnFacComments
        '
        Me.btnFacComments.Location = New System.Drawing.Point(608, 232)
        Me.btnFacComments.Name = "btnFacComments"
        Me.btnFacComments.Size = New System.Drawing.Size(80, 26)
        Me.btnFacComments.TabIndex = 1050
        Me.btnFacComments.Text = "Comments"
        '
        'btnFacFlags
        '
        Me.btnFacFlags.Location = New System.Drawing.Point(528, 232)
        Me.btnFacFlags.Name = "btnFacFlags"
        Me.btnFacFlags.Size = New System.Drawing.Size(75, 26)
        Me.btnFacFlags.TabIndex = 1051
        Me.btnFacFlags.Text = "Flags"
        '
        'txtDueByNF
        '
        Me.txtDueByNF.AcceptsTab = True
        Me.txtDueByNF.AutoSize = False
        Me.txtDueByNF.Enabled = False
        Me.txtDueByNF.Location = New System.Drawing.Point(944, 136)
        Me.txtDueByNF.Name = "txtDueByNF"
        Me.txtDueByNF.Size = New System.Drawing.Size(64, 21)
        Me.txtDueByNF.TabIndex = 1049
        Me.txtDueByNF.Text = ""
        Me.txtDueByNF.Visible = False
        Me.txtDueByNF.WordWrap = False
        '
        'lblFacilitySigNFDue
        '
        Me.lblFacilitySigNFDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySigNFDue.ForeColor = System.Drawing.Color.Red
        Me.lblFacilitySigNFDue.Location = New System.Drawing.Point(960, 112)
        Me.lblFacilitySigNFDue.Name = "lblFacilitySigNFDue"
        Me.lblFacilitySigNFDue.Size = New System.Drawing.Size(112, 23)
        Me.lblFacilitySigNFDue.TabIndex = 1048
        Me.lblFacilitySigNFDue.Text = "Due By:"
        Me.lblFacilitySigNFDue.Visible = False
        '
        'lblDateTransfered
        '
        Me.lblDateTransfered.BackColor = System.Drawing.SystemColors.Control
        Me.lblDateTransfered.Enabled = False
        Me.lblDateTransfered.Location = New System.Drawing.Point(936, 64)
        Me.lblDateTransfered.Name = "lblDateTransfered"
        Me.lblDateTransfered.TabIndex = 1047
        Me.lblDateTransfered.Visible = False
        '
        'btnFacilityCancel
        '
        Me.btnFacilityCancel.Enabled = False
        Me.btnFacilityCancel.Location = New System.Drawing.Point(448, 232)
        Me.btnFacilityCancel.Name = "btnFacilityCancel"
        Me.btnFacilityCancel.Size = New System.Drawing.Size(75, 26)
        Me.btnFacilityCancel.TabIndex = 1046
        Me.btnFacilityCancel.Text = "Cancel"
        '
        'btnFacilitySave
        '
        Me.btnFacilitySave.Enabled = False
        Me.btnFacilitySave.Location = New System.Drawing.Point(344, 232)
        Me.btnFacilitySave.Name = "btnFacilitySave"
        Me.btnFacilitySave.Size = New System.Drawing.Size(96, 26)
        Me.btnFacilitySave.TabIndex = 1045
        Me.btnFacilitySave.Text = "Save Facility"
        Me.btnFacilitySave.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnTransferLUSTEvent
        '
        Me.btnTransferLUSTEvent.Location = New System.Drawing.Point(696, 232)
        Me.btnTransferLUSTEvent.Name = "btnTransferLUSTEvent"
        Me.btnTransferLUSTEvent.Size = New System.Drawing.Size(136, 26)
        Me.btnTransferLUSTEvent.TabIndex = 24
        Me.btnTransferLUSTEvent.Text = "Transfer LUST Event"
        '
        'dtPickUpcomingInstallDateValue
        '
        Me.dtPickUpcomingInstallDateValue.Checked = False
        Me.dtPickUpcomingInstallDateValue.Enabled = False
        Me.dtPickUpcomingInstallDateValue.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickUpcomingInstallDateValue.Location = New System.Drawing.Point(504, 200)
        Me.dtPickUpcomingInstallDateValue.Name = "dtPickUpcomingInstallDateValue"
        Me.dtPickUpcomingInstallDateValue.Size = New System.Drawing.Size(88, 21)
        Me.dtPickUpcomingInstallDateValue.TabIndex = 12
        '
        'lblUpcomingInstallDate
        '
        Me.lblUpcomingInstallDate.Location = New System.Drawing.Point(344, 200)
        Me.lblUpcomingInstallDate.Name = "lblUpcomingInstallDate"
        Me.lblUpcomingInstallDate.Size = New System.Drawing.Size(160, 20)
        Me.lblUpcomingInstallDate.TabIndex = 1044
        Me.lblUpcomingInstallDate.Text = "Upcoming Installation Date:"
        Me.lblUpcomingInstallDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkUpcomingInstall
        '
        Me.chkUpcomingInstall.Enabled = False
        Me.chkUpcomingInstall.Location = New System.Drawing.Point(344, 176)
        Me.chkUpcomingInstall.Name = "chkUpcomingInstall"
        Me.chkUpcomingInstall.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkUpcomingInstall.Size = New System.Drawing.Size(144, 20)
        Me.chkUpcomingInstall.TabIndex = 11
        Me.chkUpcomingInstall.Text = "Upcoming Installation "
        Me.chkUpcomingInstall.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lnkLblNextFac
        '
        Me.lnkLblNextFac.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnkLblNextFac.Location = New System.Drawing.Point(256, 8)
        Me.lnkLblNextFac.Name = "lnkLblNextFac"
        Me.lnkLblNextFac.Size = New System.Drawing.Size(56, 16)
        Me.lnkLblNextFac.TabIndex = 26
        Me.lnkLblNextFac.TabStop = True
        Me.lnkLblNextFac.Text = "Next>>"
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
        Me.lblCAPStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFuelBrand
        '
        Me.txtFuelBrand.Location = New System.Drawing.Point(456, 152)
        Me.txtFuelBrand.Name = "txtFuelBrand"
        Me.txtFuelBrand.ReadOnly = True
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
        'dtFacilityPowerOff
        '
        Me.dtFacilityPowerOff.Checked = False
        Me.dtFacilityPowerOff.Enabled = False
        Me.dtFacilityPowerOff.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtFacilityPowerOff.Location = New System.Drawing.Point(1008, 232)
        Me.dtFacilityPowerOff.Name = "dtFacilityPowerOff"
        Me.dtFacilityPowerOff.ShowCheckBox = True
        Me.dtFacilityPowerOff.Size = New System.Drawing.Size(104, 21)
        Me.dtFacilityPowerOff.TabIndex = 9
        Me.dtFacilityPowerOff.Visible = False
        '
        'lnkLblPrevFacility
        '
        Me.lnkLblPrevFacility.AutoSize = True
        Me.lnkLblPrevFacility.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnkLblPrevFacility.Location = New System.Drawing.Point(176, 8)
        Me.lnkLblPrevFacility.Name = "lnkLblPrevFacility"
        Me.lnkLblPrevFacility.Size = New System.Drawing.Size(70, 17)
        Me.lnkLblPrevFacility.TabIndex = 25
        Me.lnkLblPrevFacility.TabStop = True
        Me.lnkLblPrevFacility.Text = "<< Previous"
        '
        'lblLUSTSite
        '
        Me.lblLUSTSite.Location = New System.Drawing.Point(344, 56)
        Me.lblLUSTSite.Name = "lblLUSTSite"
        Me.lblLUSTSite.Size = New System.Drawing.Size(104, 20)
        Me.lblLUSTSite.TabIndex = 1030
        Me.lblLUSTSite.Text = "Active LUST Site:"
        Me.lblLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkLUSTSite
        '
        Me.chkLUSTSite.Enabled = False
        Me.chkLUSTSite.Location = New System.Drawing.Point(456, 56)
        Me.chkLUSTSite.Name = "chkLUSTSite"
        Me.chkLUSTSite.Size = New System.Drawing.Size(16, 16)
        Me.chkLUSTSite.TabIndex = 6
        Me.chkLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPowerOff
        '
        Me.lblPowerOff.Location = New System.Drawing.Point(920, 232)
        Me.lblPowerOff.Name = "lblPowerOff"
        Me.lblPowerOff.Size = New System.Drawing.Size(80, 23)
        Me.lblPowerOff.TabIndex = 1028
        Me.lblPowerOff.Text = "Power Off"
        Me.lblPowerOff.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblPowerOff.Visible = False
        '
        'lblCAPCandidate
        '
        Me.lblCAPCandidate.Location = New System.Drawing.Point(344, 104)
        Me.lblCAPCandidate.Name = "lblCAPCandidate"
        Me.lblCAPCandidate.Size = New System.Drawing.Size(100, 20)
        Me.lblCAPCandidate.TabIndex = 1026
        Me.lblCAPCandidate.Text = "CAP Candidate:"
        Me.lblCAPCandidate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkCAPCandidate
        '
        Me.chkCAPCandidate.Enabled = False
        Me.chkCAPCandidate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCAPCandidate.Location = New System.Drawing.Point(456, 104)
        Me.chkCAPCandidate.Name = "chkCAPCandidate"
        Me.chkCAPCandidate.Size = New System.Drawing.Size(16, 20)
        Me.chkCAPCandidate.TabIndex = 7
        Me.chkCAPCandidate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFacilityLocationType
        '
        Me.lblFacilityLocationType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLocationType.Location = New System.Drawing.Point(664, 152)
        Me.lblFacilityLocationType.Name = "lblFacilityLocationType"
        Me.lblFacilityLocationType.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityLocationType.TabIndex = 1024
        Me.lblFacilityLocationType.Text = "Type:"
        '
        'cmbFacilityLocationType
        '
        Me.cmbFacilityLocationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityLocationType.DropDownWidth = 250
        Me.cmbFacilityLocationType.Enabled = False
        Me.cmbFacilityLocationType.ItemHeight = 15
        Me.cmbFacilityLocationType.Location = New System.Drawing.Point(752, 152)
        Me.cmbFacilityLocationType.Name = "cmbFacilityLocationType"
        Me.cmbFacilityLocationType.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityLocationType.TabIndex = 23
        '
        'lblFacilityMethod
        '
        Me.lblFacilityMethod.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityMethod.Location = New System.Drawing.Point(664, 128)
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
        Me.cmbFacilityMethod.Location = New System.Drawing.Point(752, 128)
        Me.cmbFacilityMethod.Name = "cmbFacilityMethod"
        Me.cmbFacilityMethod.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityMethod.TabIndex = 22
        '
        'lblFacilityDatum
        '
        Me.lblFacilityDatum.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityDatum.Location = New System.Drawing.Point(664, 104)
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
        Me.cmbFacilityDatum.Location = New System.Drawing.Point(752, 104)
        Me.cmbFacilityDatum.Name = "cmbFacilityDatum"
        Me.cmbFacilityDatum.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityDatum.TabIndex = 21
        '
        'cmbFacilityType
        '
        Me.cmbFacilityType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityType.DropDownWidth = 180
        Me.cmbFacilityType.Enabled = False
        Me.cmbFacilityType.ItemHeight = 15
        Me.cmbFacilityType.Location = New System.Drawing.Point(752, 32)
        Me.cmbFacilityType.Name = "cmbFacilityType"
        Me.cmbFacilityType.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityType.TabIndex = 14
        '
        'txtFacilityLatSec
        '
        Me.txtFacilityLatSec.AcceptsTab = True
        Me.txtFacilityLatSec.AutoSize = False
        Me.txtFacilityLatSec.Location = New System.Drawing.Point(848, 56)
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
        Me.txtFacilityLongSec.Location = New System.Drawing.Point(848, 80)
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
        Me.txtFacilityLatMin.Location = New System.Drawing.Point(800, 56)
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
        Me.txtFacilityLongMin.Location = New System.Drawing.Point(800, 80)
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
        Me.lblFacilityLongMin.Location = New System.Drawing.Point(832, 80)
        Me.lblFacilityLongMin.Name = "lblFacilityLongMin"
        Me.lblFacilityLongMin.Size = New System.Drawing.Size(8, 16)
        Me.lblFacilityLongMin.TabIndex = 1018
        Me.lblFacilityLongMin.Text = "'"
        '
        'lblFacilityLongSec
        '
        Me.lblFacilityLongSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongSec.Location = New System.Drawing.Point(888, 80)
        Me.lblFacilityLongSec.Name = "lblFacilityLongSec"
        Me.lblFacilityLongSec.Size = New System.Drawing.Size(10, 16)
        Me.lblFacilityLongSec.TabIndex = 1017
        Me.lblFacilityLongSec.Text = """"
        '
        'lblFacilityLatMin
        '
        Me.lblFacilityLatMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatMin.Location = New System.Drawing.Point(832, 56)
        Me.lblFacilityLatMin.Name = "lblFacilityLatMin"
        Me.lblFacilityLatMin.Size = New System.Drawing.Size(8, 16)
        Me.lblFacilityLatMin.TabIndex = 1016
        Me.lblFacilityLatMin.Text = "'"
        '
        'lblFacilityLatSec
        '
        Me.lblFacilityLatSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatSec.Location = New System.Drawing.Point(888, 56)
        Me.lblFacilityLatSec.Name = "lblFacilityLatSec"
        Me.lblFacilityLatSec.Size = New System.Drawing.Size(10, 16)
        Me.lblFacilityLatSec.TabIndex = 1015
        Me.lblFacilityLatSec.Text = """"
        '
        'lblFacilityLongDegree
        '
        Me.lblFacilityLongDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongDegree.Location = New System.Drawing.Point(784, 76)
        Me.lblFacilityLongDegree.Name = "lblFacilityLongDegree"
        Me.lblFacilityLongDegree.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLongDegree.TabIndex = 1010
        Me.lblFacilityLongDegree.Text = "o"
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
        Me.lblFacilitySIC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        Me.dtPickFacilityRecvd.Enabled = False
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
        Me.lblDateReceived.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkSignatureofNF
        '
        Me.chkSignatureofNF.Enabled = False
        Me.chkSignatureofNF.Location = New System.Drawing.Point(136, 217)
        Me.chkSignatureofNF.Name = "chkSignatureofNF"
        Me.chkSignatureofNF.Size = New System.Drawing.Size(16, 16)
        Me.chkSignatureofNF.TabIndex = 4
        Me.chkSignatureofNF.Text = "CheckBox5"
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
        Me.lblFacilityFuelBrand.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityFuelBrand.TabIndex = 125
        Me.lblFacilityFuelBrand.Text = "Fuel Brand:"
        Me.lblFacilityFuelBrand.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFacilityStatusValue
        '
        Me.lblFacilityStatusValue.BackColor = System.Drawing.Color.Transparent
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
        Me.lblFacilityStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFacilityLongDegree
        '
        Me.txtFacilityLongDegree.AcceptsTab = True
        Me.txtFacilityLongDegree.AutoSize = False
        Me.txtFacilityLongDegree.Location = New System.Drawing.Point(752, 80)
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
        Me.txtFacilityLatDegree.Location = New System.Drawing.Point(752, 56)
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
        Me.lblFacilityLongitude.Location = New System.Drawing.Point(664, 80)
        Me.lblFacilityLongitude.Name = "lblFacilityLongitude"
        Me.lblFacilityLongitude.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityLongitude.TabIndex = 111
        Me.lblFacilityLongitude.Text = "Longitude:"
        '
        'lblFacilityLatitude
        '
        Me.lblFacilityLatitude.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatitude.Location = New System.Drawing.Point(664, 56)
        Me.lblFacilityLatitude.Name = "lblFacilityLatitude"
        Me.lblFacilityLatitude.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityLatitude.TabIndex = 110
        Me.lblFacilityLatitude.Text = "Latitude:"
        '
        'lblFacilityType
        '
        Me.lblFacilityType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityType.Location = New System.Drawing.Point(664, 32)
        Me.lblFacilityType.Name = "lblFacilityType"
        Me.lblFacilityType.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityType.TabIndex = 106
        Me.lblFacilityType.Text = "Facility Type:"
        '
        'txtFacilityAIID
        '
        Me.txtFacilityAIID.AcceptsTab = True
        Me.txtFacilityAIID.AutoSize = False
        Me.txtFacilityAIID.Enabled = False
        Me.txtFacilityAIID.Location = New System.Drawing.Point(752, 8)
        Me.txtFacilityAIID.Name = "txtFacilityAIID"
        Me.txtFacilityAIID.ReadOnly = True
        Me.txtFacilityAIID.Size = New System.Drawing.Size(136, 20)
        Me.txtFacilityAIID.TabIndex = 13
        Me.txtFacilityAIID.Text = ""
        Me.txtFacilityAIID.WordWrap = False
        '
        'lblfacilityAIID
        '
        Me.lblfacilityAIID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblfacilityAIID.Location = New System.Drawing.Point(664, 8)
        Me.lblfacilityAIID.Name = "lblfacilityAIID"
        Me.lblfacilityAIID.Size = New System.Drawing.Size(80, 20)
        Me.lblfacilityAIID.TabIndex = 104
        Me.lblfacilityAIID.Text = "Facility AIID:"
        '
        'lblFacilityIDValue
        '
        Me.lblFacilityIDValue.Enabled = False
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
        Me.txtFacilityName.ReadOnly = True
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
        'lblFacilityLatDegree
        '
        Me.lblFacilityLatDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatDegree.Location = New System.Drawing.Point(784, 50)
        Me.lblFacilityLatDegree.Name = "lblFacilityLatDegree"
        Me.lblFacilityLatDegree.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLatDegree.TabIndex = 1009
        Me.lblFacilityLatDegree.Text = "o"
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
        'tbPageLUSTEvent
        '
        Me.tbPageLUSTEvent.Controls.Add(Me.pnlLustEvents)
        Me.tbPageLUSTEvent.Controls.Add(Me.pnlLUSTEventHeader)
        Me.tbPageLUSTEvent.Controls.Add(Me.pnlLUSTEventBottom)
        Me.tbPageLUSTEvent.Location = New System.Drawing.Point(4, 22)
        Me.tbPageLUSTEvent.Name = "tbPageLUSTEvent"
        Me.tbPageLUSTEvent.Size = New System.Drawing.Size(1016, 644)
        Me.tbPageLUSTEvent.TabIndex = 12
        Me.tbPageLUSTEvent.Text = "LUST Events"
        Me.tbPageLUSTEvent.Visible = False
        '
        'pnlLustEvents
        '
        Me.pnlLustEvents.AutoScroll = True
        Me.pnlLustEvents.Controls.Add(Me.pnlLUSTEventDetails)
        Me.pnlLustEvents.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlLustEvents.Location = New System.Drawing.Point(0, 56)
        Me.pnlLustEvents.Name = "pnlLustEvents"
        Me.pnlLustEvents.Size = New System.Drawing.Size(1016, 548)
        Me.pnlLustEvents.TabIndex = 6
        '
        'pnlLUSTEventDetails
        '
        Me.pnlLUSTEventDetails.AutoScroll = True
        Me.pnlLUSTEventDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlERACandContactsDetails)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlERACandContacts)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlRemediationSystemsDetails)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlRemediationSystems)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlCommentsDetails)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlComments)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlActivitiesDocumentsDetails)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlActivitiesDocuments)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlFundsEligibilityDetails)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlFundsEligibility)
        Me.pnlLUSTEventDetails.Controls.Add(Me.PnlReleaseInfoDetails)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlReleaseInfo)
        Me.pnlLUSTEventDetails.Controls.Add(Me.PnlEventInfoDetails)
        Me.pnlLUSTEventDetails.Controls.Add(Me.pnlEventInfo)
        Me.pnlLUSTEventDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlLUSTEventDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlLUSTEventDetails.Name = "pnlLUSTEventDetails"
        Me.pnlLUSTEventDetails.Size = New System.Drawing.Size(1016, 548)
        Me.pnlLUSTEventDetails.TabIndex = 0
        '
        'pnlERACandContactsDetails
        '
        Me.pnlERACandContactsDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlERACandContactsDetails.Controls.Add(Me.pnlERACContactButtons)
        Me.pnlERACandContactsDetails.Controls.Add(Me.pnlERACContactContainer)
        Me.pnlERACandContactsDetails.Controls.Add(Me.pnlERACContactHeader)
        Me.pnlERACandContactsDetails.Controls.Add(Me.pnlERACandIRAC)
        Me.pnlERACandContactsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlERACandContactsDetails.Location = New System.Drawing.Point(0, 1792)
        Me.pnlERACandContactsDetails.Name = "pnlERACandContactsDetails"
        Me.pnlERACandContactsDetails.Size = New System.Drawing.Size(996, 360)
        Me.pnlERACandContactsDetails.TabIndex = 202
        '
        'pnlERACContactButtons
        '
        Me.pnlERACContactButtons.Controls.Add(Me.btnERACContactModify)
        Me.pnlERACContactButtons.Controls.Add(Me.btnERACContactDelete)
        Me.pnlERACContactButtons.Controls.Add(Me.btnERACContactAssociate)
        Me.pnlERACContactButtons.Controls.Add(Me.btnERACContactAddorSearch)
        Me.pnlERACContactButtons.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlERACContactButtons.DockPadding.All = 3
        Me.pnlERACContactButtons.Location = New System.Drawing.Point(0, 312)
        Me.pnlERACContactButtons.Name = "pnlERACContactButtons"
        Me.pnlERACContactButtons.Size = New System.Drawing.Size(994, 48)
        Me.pnlERACContactButtons.TabIndex = 3
        '
        'btnERACContactModify
        '
        Me.btnERACContactModify.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnERACContactModify.Location = New System.Drawing.Point(240, 13)
        Me.btnERACContactModify.Name = "btnERACContactModify"
        Me.btnERACContactModify.Size = New System.Drawing.Size(235, 26)
        Me.btnERACContactModify.TabIndex = 1
        Me.btnERACContactModify.Text = "Modify Contact"
        '
        'btnERACContactDelete
        '
        Me.btnERACContactDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnERACContactDelete.Location = New System.Drawing.Point(472, 13)
        Me.btnERACContactDelete.Name = "btnERACContactDelete"
        Me.btnERACContactDelete.Size = New System.Drawing.Size(235, 26)
        Me.btnERACContactDelete.TabIndex = 2
        Me.btnERACContactDelete.Text = "Disassociate Contact"
        '
        'btnERACContactAssociate
        '
        Me.btnERACContactAssociate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnERACContactAssociate.Location = New System.Drawing.Point(704, 13)
        Me.btnERACContactAssociate.Name = "btnERACContactAssociate"
        Me.btnERACContactAssociate.Size = New System.Drawing.Size(235, 26)
        Me.btnERACContactAssociate.TabIndex = 3
        Me.btnERACContactAssociate.Text = "Associate Contact from Different Module"
        '
        'btnERACContactAddorSearch
        '
        Me.btnERACContactAddorSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnERACContactAddorSearch.Location = New System.Drawing.Point(8, 13)
        Me.btnERACContactAddorSearch.Name = "btnERACContactAddorSearch"
        Me.btnERACContactAddorSearch.Size = New System.Drawing.Size(235, 26)
        Me.btnERACContactAddorSearch.TabIndex = 0
        Me.btnERACContactAddorSearch.Text = "Add/Search Contact to Associate"
        '
        'pnlERACContactContainer
        '
        Me.pnlERACContactContainer.AutoScroll = True
        Me.pnlERACContactContainer.Controls.Add(Me.ugERACContacts)
        Me.pnlERACContactContainer.Controls.Add(Me.Label4)
        Me.pnlERACContactContainer.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlERACContactContainer.Location = New System.Drawing.Point(0, 80)
        Me.pnlERACContactContainer.Name = "pnlERACContactContainer"
        Me.pnlERACContactContainer.Size = New System.Drawing.Size(994, 232)
        Me.pnlERACContactContainer.TabIndex = 2
        '
        'ugERACContacts
        '
        Me.ugERACContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugERACContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugERACContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugERACContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugERACContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugERACContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugERACContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugERACContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugERACContacts.Name = "ugERACContacts"
        Me.ugERACContacts.Size = New System.Drawing.Size(994, 232)
        Me.ugERACContacts.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(792, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(7, 23)
        Me.Label4.TabIndex = 2
        '
        'pnlERACContactHeader
        '
        Me.pnlERACContactHeader.Controls.Add(Me.chkERACShowActive)
        Me.pnlERACContactHeader.Controls.Add(Me.chkERACShowRelatedContacts)
        Me.pnlERACContactHeader.Controls.Add(Me.chkERACShowContactsforAllModules)
        Me.pnlERACContactHeader.Controls.Add(Me.lblLustContacts)
        Me.pnlERACContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlERACContactHeader.DockPadding.All = 3
        Me.pnlERACContactHeader.Location = New System.Drawing.Point(0, 48)
        Me.pnlERACContactHeader.Name = "pnlERACContactHeader"
        Me.pnlERACContactHeader.Size = New System.Drawing.Size(994, 32)
        Me.pnlERACContactHeader.TabIndex = 1
        '
        'chkERACShowActive
        '
        Me.chkERACShowActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkERACShowActive.Location = New System.Drawing.Point(635, 8)
        Me.chkERACShowActive.Name = "chkERACShowActive"
        Me.chkERACShowActive.Size = New System.Drawing.Size(144, 16)
        Me.chkERACShowActive.TabIndex = 2
        Me.chkERACShowActive.Tag = "646"
        Me.chkERACShowActive.Text = "Show Active Only"
        '
        'chkERACShowRelatedContacts
        '
        Me.chkERACShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkERACShowRelatedContacts.Location = New System.Drawing.Point(467, 8)
        Me.chkERACShowRelatedContacts.Name = "chkERACShowRelatedContacts"
        Me.chkERACShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkERACShowRelatedContacts.TabIndex = 1
        Me.chkERACShowRelatedContacts.Tag = "645"
        Me.chkERACShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkERACShowContactsforAllModules
        '
        Me.chkERACShowContactsforAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkERACShowContactsforAllModules.Location = New System.Drawing.Point(251, 8)
        Me.chkERACShowContactsforAllModules.Name = "chkERACShowContactsforAllModules"
        Me.chkERACShowContactsforAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkERACShowContactsforAllModules.TabIndex = 0
        Me.chkERACShowContactsforAllModules.Tag = "644"
        Me.chkERACShowContactsforAllModules.Text = "Show Contacts for All Modules"
        '
        'lblLustContacts
        '
        Me.lblLustContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLustContacts.Location = New System.Drawing.Point(8, 8)
        Me.lblLustContacts.Name = "lblLustContacts"
        Me.lblLustContacts.Size = New System.Drawing.Size(104, 16)
        Me.lblLustContacts.TabIndex = 139
        Me.lblLustContacts.Text = "LUST Contacts"
        '
        'pnlERACandIRAC
        '
        Me.pnlERACandIRAC.Controls.Add(Me.BtnSaveEngineers)
        Me.pnlERACandIRAC.Controls.Add(Me.txtERAC)
        Me.pnlERACandIRAC.Controls.Add(Me.txtIRAC)
        Me.pnlERACandIRAC.Controls.Add(Me.lblIRACSearch)
        Me.pnlERACandIRAC.Controls.Add(Me.lblERACSearch)
        Me.pnlERACandIRAC.Controls.Add(Me.lblERAC)
        Me.pnlERACandIRAC.Controls.Add(Me.lblIRAC)
        Me.pnlERACandIRAC.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlERACandIRAC.DockPadding.All = 3
        Me.pnlERACandIRAC.Location = New System.Drawing.Point(0, 0)
        Me.pnlERACandIRAC.Name = "pnlERACandIRAC"
        Me.pnlERACandIRAC.Size = New System.Drawing.Size(994, 48)
        Me.pnlERACandIRAC.TabIndex = 0
        '
        'BtnSaveEngineers
        '
        Me.BtnSaveEngineers.Enabled = False
        Me.BtnSaveEngineers.Location = New System.Drawing.Point(672, 16)
        Me.BtnSaveEngineers.Name = "BtnSaveEngineers"
        Me.BtnSaveEngineers.Size = New System.Drawing.Size(160, 23)
        Me.BtnSaveEngineers.TabIndex = 224
        Me.BtnSaveEngineers.Text = "Save ERAC/IRAC Contacts"
        '
        'txtERAC
        '
        Me.txtERAC.Location = New System.Drawing.Point(168, 16)
        Me.txtERAC.Name = "txtERAC"
        Me.txtERAC.Size = New System.Drawing.Size(184, 21)
        Me.txtERAC.TabIndex = 223
        Me.txtERAC.Text = ""
        '
        'txtIRAC
        '
        Me.txtIRAC.Location = New System.Drawing.Point(432, 16)
        Me.txtIRAC.Name = "txtIRAC"
        Me.txtIRAC.Size = New System.Drawing.Size(184, 21)
        Me.txtIRAC.TabIndex = 222
        Me.txtIRAC.Text = ""
        '
        'lblIRACSearch
        '
        Me.lblIRACSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblIRACSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIRACSearch.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblIRACSearch.Location = New System.Drawing.Point(632, 16)
        Me.lblIRACSearch.Name = "lblIRACSearch"
        Me.lblIRACSearch.Size = New System.Drawing.Size(16, 22)
        Me.lblIRACSearch.TabIndex = 221
        Me.lblIRACSearch.Text = "?"
        Me.lblIRACSearch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblERACSearch
        '
        Me.lblERACSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblERACSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblERACSearch.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblERACSearch.Location = New System.Drawing.Point(360, 16)
        Me.lblERACSearch.Name = "lblERACSearch"
        Me.lblERACSearch.Size = New System.Drawing.Size(16, 22)
        Me.lblERACSearch.TabIndex = 220
        Me.lblERACSearch.Text = "?"
        Me.lblERACSearch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblERAC
        '
        Me.lblERAC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblERAC.Location = New System.Drawing.Point(16, 16)
        Me.lblERAC.Name = "lblERAC"
        Me.lblERAC.Size = New System.Drawing.Size(144, 16)
        Me.lblERAC.TabIndex = 185
        Me.lblERAC.Text = "Environmental Contact:"
        Me.lblERAC.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblIRAC
        '
        Me.lblIRAC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIRAC.Location = New System.Drawing.Point(384, 16)
        Me.lblIRAC.Name = "lblIRAC"
        Me.lblIRAC.Size = New System.Drawing.Size(40, 16)
        Me.lblIRAC.TabIndex = 185
        Me.lblIRAC.Text = "IRAC:"
        Me.lblIRAC.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'pnlERACandContacts
        '
        Me.pnlERACandContacts.Controls.Add(Me.lblERACandContactsHead)
        Me.pnlERACandContacts.Controls.Add(Me.lblERACandContactsDisplay)
        Me.pnlERACandContacts.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlERACandContacts.Location = New System.Drawing.Point(0, 1768)
        Me.pnlERACandContacts.Name = "pnlERACandContacts"
        Me.pnlERACandContacts.Size = New System.Drawing.Size(996, 24)
        Me.pnlERACandContacts.TabIndex = 201
        '
        'lblERACandContactsHead
        '
        Me.lblERACandContactsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblERACandContactsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblERACandContactsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblERACandContactsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblERACandContactsHead.Name = "lblERACandContactsHead"
        Me.lblERACandContactsHead.Size = New System.Drawing.Size(980, 24)
        Me.lblERACandContactsHead.TabIndex = 1
        Me.lblERACandContactsHead.Text = "Environmental Contact / Contacts"
        Me.lblERACandContactsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblERACandContactsDisplay
        '
        Me.lblERACandContactsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblERACandContactsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblERACandContactsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblERACandContactsDisplay.Name = "lblERACandContactsDisplay"
        Me.lblERACandContactsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblERACandContactsDisplay.TabIndex = 0
        Me.lblERACandContactsDisplay.Text = "-"
        Me.lblERACandContactsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlRemediationSystemsDetails
        '
        Me.pnlRemediationSystemsDetails.AutoScroll = True
        Me.pnlRemediationSystemsDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlRemediationSystemsDetails.Controls.Add(Me.btnAddRemediationSystem)
        Me.pnlRemediationSystemsDetails.Controls.Add(Me.ugRemediationSystem)
        Me.pnlRemediationSystemsDetails.Controls.Add(Me.btnModifyRemediationSystem)
        Me.pnlRemediationSystemsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlRemediationSystemsDetails.Location = New System.Drawing.Point(0, 1544)
        Me.pnlRemediationSystemsDetails.Name = "pnlRemediationSystemsDetails"
        Me.pnlRemediationSystemsDetails.Size = New System.Drawing.Size(996, 224)
        Me.pnlRemediationSystemsDetails.TabIndex = 200
        '
        'btnAddRemediationSystem
        '
        Me.btnAddRemediationSystem.Location = New System.Drawing.Point(216, 192)
        Me.btnAddRemediationSystem.Name = "btnAddRemediationSystem"
        Me.btnAddRemediationSystem.Size = New System.Drawing.Size(176, 23)
        Me.btnAddRemediationSystem.TabIndex = 2
        Me.btnAddRemediationSystem.Text = "Add Remediation System"
        '
        'ugRemediationSystem
        '
        Me.ugRemediationSystem.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugRemediationSystem.DisplayLayout.AutoFitColumns = True
        Me.ugRemediationSystem.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugRemediationSystem.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugRemediationSystem.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugRemediationSystem.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugRemediationSystem.Location = New System.Drawing.Point(24, 8)
        Me.ugRemediationSystem.Name = "ugRemediationSystem"
        Me.ugRemediationSystem.Size = New System.Drawing.Size(768, 176)
        Me.ugRemediationSystem.TabIndex = 1
        '
        'btnModifyRemediationSystem
        '
        Me.btnModifyRemediationSystem.Location = New System.Drawing.Point(24, 192)
        Me.btnModifyRemediationSystem.Name = "btnModifyRemediationSystem"
        Me.btnModifyRemediationSystem.Size = New System.Drawing.Size(176, 23)
        Me.btnModifyRemediationSystem.TabIndex = 0
        Me.btnModifyRemediationSystem.Text = "Modify Remediation System"
        '
        'pnlRemediationSystems
        '
        Me.pnlRemediationSystems.Controls.Add(Me.lblRemediationSystemsHead)
        Me.pnlRemediationSystems.Controls.Add(Me.lblRemediationSystemsDisplay)
        Me.pnlRemediationSystems.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlRemediationSystems.Location = New System.Drawing.Point(0, 1520)
        Me.pnlRemediationSystems.Name = "pnlRemediationSystems"
        Me.pnlRemediationSystems.Size = New System.Drawing.Size(996, 24)
        Me.pnlRemediationSystems.TabIndex = 199
        '
        'lblRemediationSystemsHead
        '
        Me.lblRemediationSystemsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblRemediationSystemsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblRemediationSystemsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblRemediationSystemsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblRemediationSystemsHead.Name = "lblRemediationSystemsHead"
        Me.lblRemediationSystemsHead.Size = New System.Drawing.Size(980, 24)
        Me.lblRemediationSystemsHead.TabIndex = 1
        Me.lblRemediationSystemsHead.Text = "Remediation Systems"
        Me.lblRemediationSystemsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblRemediationSystemsDisplay
        '
        Me.lblRemediationSystemsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblRemediationSystemsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblRemediationSystemsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblRemediationSystemsDisplay.Name = "lblRemediationSystemsDisplay"
        Me.lblRemediationSystemsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblRemediationSystemsDisplay.TabIndex = 0
        Me.lblRemediationSystemsDisplay.Text = "-"
        Me.lblRemediationSystemsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlCommentsDetails
        '
        Me.pnlCommentsDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCommentsDetails.Controls.Add(Me.btnViewModifyComment)
        Me.pnlCommentsDetails.Controls.Add(Me.btnFlagsLustEvent)
        Me.pnlCommentsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCommentsDetails.Location = New System.Drawing.Point(0, 1472)
        Me.pnlCommentsDetails.Name = "pnlCommentsDetails"
        Me.pnlCommentsDetails.Size = New System.Drawing.Size(996, 48)
        Me.pnlCommentsDetails.TabIndex = 197
        '
        'btnViewModifyComment
        '
        Me.btnViewModifyComment.Location = New System.Drawing.Point(48, 8)
        Me.btnViewModifyComment.Name = "btnViewModifyComment"
        Me.btnViewModifyComment.Size = New System.Drawing.Size(96, 23)
        Me.btnViewModifyComment.TabIndex = 2
        Me.btnViewModifyComment.Text = "Comments"
        '
        'btnFlagsLustEvent
        '
        Me.btnFlagsLustEvent.Location = New System.Drawing.Point(160, 8)
        Me.btnFlagsLustEvent.Name = "btnFlagsLustEvent"
        Me.btnFlagsLustEvent.Size = New System.Drawing.Size(112, 23)
        Me.btnFlagsLustEvent.TabIndex = 7
        Me.btnFlagsLustEvent.Text = "Flags"
        '
        'pnlComments
        '
        Me.pnlComments.Controls.Add(Me.lblCommentsHead)
        Me.pnlComments.Controls.Add(Me.lblCommentsDisplay)
        Me.pnlComments.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlComments.Location = New System.Drawing.Point(0, 1448)
        Me.pnlComments.Name = "pnlComments"
        Me.pnlComments.Size = New System.Drawing.Size(996, 24)
        Me.pnlComments.TabIndex = 196
        '
        'lblCommentsHead
        '
        Me.lblCommentsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblCommentsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCommentsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCommentsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblCommentsHead.Name = "lblCommentsHead"
        Me.lblCommentsHead.Size = New System.Drawing.Size(980, 24)
        Me.lblCommentsHead.TabIndex = 1
        Me.lblCommentsHead.Text = "Comments / Flags"
        Me.lblCommentsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCommentsDisplay
        '
        Me.lblCommentsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCommentsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCommentsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblCommentsDisplay.Name = "lblCommentsDisplay"
        Me.lblCommentsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblCommentsDisplay.TabIndex = 0
        Me.lblCommentsDisplay.Text = "-"
        Me.lblCommentsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlActivitiesDocumentsDetails
        '
        Me.pnlActivitiesDocumentsDetails.AutoScroll = True
        Me.pnlActivitiesDocumentsDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.lblActivitiesDocumentsCellDesc)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.btnDeleteActivity)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.btnDeleteDocument)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.btnModifyDocument)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.btnAddDocument)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.btnModifyActivity)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.btnAddActivity)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.ugActivitiesandDocuments)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.chkShowallDocuments)
        Me.pnlActivitiesDocumentsDetails.Controls.Add(Me.chkShowallActivities)
        Me.pnlActivitiesDocumentsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlActivitiesDocumentsDetails.Location = New System.Drawing.Point(0, 1088)
        Me.pnlActivitiesDocumentsDetails.Name = "pnlActivitiesDocumentsDetails"
        Me.pnlActivitiesDocumentsDetails.Size = New System.Drawing.Size(996, 360)
        Me.pnlActivitiesDocumentsDetails.TabIndex = 195
        '
        'lblActivitiesDocumentsCellDesc
        '
        Me.lblActivitiesDocumentsCellDesc.Location = New System.Drawing.Point(456, 2)
        Me.lblActivitiesDocumentsCellDesc.Name = "lblActivitiesDocumentsCellDesc"
        Me.lblActivitiesDocumentsCellDesc.Size = New System.Drawing.Size(360, 32)
        Me.lblActivitiesDocumentsCellDesc.TabIndex = 9
        '
        'btnDeleteActivity
        '
        Me.btnDeleteActivity.Location = New System.Drawing.Point(616, 312)
        Me.btnDeleteActivity.Name = "btnDeleteActivity"
        Me.btnDeleteActivity.Size = New System.Drawing.Size(104, 23)
        Me.btnDeleteActivity.TabIndex = 7
        Me.btnDeleteActivity.Text = "Delete Activity"
        '
        'btnDeleteDocument
        '
        Me.btnDeleteDocument.Location = New System.Drawing.Point(720, 312)
        Me.btnDeleteDocument.Name = "btnDeleteDocument"
        Me.btnDeleteDocument.Size = New System.Drawing.Size(112, 23)
        Me.btnDeleteDocument.TabIndex = 8
        Me.btnDeleteDocument.Text = "Delete Document"
        '
        'btnModifyDocument
        '
        Me.btnModifyDocument.Location = New System.Drawing.Point(432, 312)
        Me.btnModifyDocument.Name = "btnModifyDocument"
        Me.btnModifyDocument.Size = New System.Drawing.Size(112, 23)
        Me.btnModifyDocument.TabIndex = 6
        Me.btnModifyDocument.Text = "Modify Document"
        '
        'btnAddDocument
        '
        Me.btnAddDocument.Location = New System.Drawing.Point(328, 312)
        Me.btnAddDocument.Name = "btnAddDocument"
        Me.btnAddDocument.Size = New System.Drawing.Size(104, 23)
        Me.btnAddDocument.TabIndex = 4
        Me.btnAddDocument.Text = "Add Document"
        '
        'btnModifyActivity
        '
        Me.btnModifyActivity.Location = New System.Drawing.Point(120, 312)
        Me.btnModifyActivity.Name = "btnModifyActivity"
        Me.btnModifyActivity.Size = New System.Drawing.Size(104, 23)
        Me.btnModifyActivity.TabIndex = 5
        Me.btnModifyActivity.Text = "Modify Activity"
        '
        'btnAddActivity
        '
        Me.btnAddActivity.Location = New System.Drawing.Point(24, 312)
        Me.btnAddActivity.Name = "btnAddActivity"
        Me.btnAddActivity.Size = New System.Drawing.Size(96, 23)
        Me.btnAddActivity.TabIndex = 3
        Me.btnAddActivity.Text = "Add Activity"
        '
        'ugActivitiesandDocuments
        '
        Me.ugActivitiesandDocuments.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugActivitiesandDocuments.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugActivitiesandDocuments.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugActivitiesandDocuments.Location = New System.Drawing.Point(25, 35)
        Me.ugActivitiesandDocuments.Name = "ugActivitiesandDocuments"
        Me.ugActivitiesandDocuments.Size = New System.Drawing.Size(800, 269)
        Me.ugActivitiesandDocuments.TabIndex = 2
        Me.ugActivitiesandDocuments.Text = "Activities / Documents"
        '
        'chkShowallDocuments
        '
        Me.chkShowallDocuments.Location = New System.Drawing.Point(160, 12)
        Me.chkShowallDocuments.Name = "chkShowallDocuments"
        Me.chkShowallDocuments.Size = New System.Drawing.Size(288, 16)
        Me.chkShowallDocuments.TabIndex = 1
        Me.chkShowallDocuments.Text = "Show All Documents for Activites Shown Below"
        '
        'chkShowallActivities
        '
        Me.chkShowallActivities.Location = New System.Drawing.Point(24, 12)
        Me.chkShowallActivities.Name = "chkShowallActivities"
        Me.chkShowallActivities.Size = New System.Drawing.Size(128, 16)
        Me.chkShowallActivities.TabIndex = 0
        Me.chkShowallActivities.Text = "Show all Activities"
        '
        'pnlActivitiesDocuments
        '
        Me.pnlActivitiesDocuments.Controls.Add(Me.lblActivitiesDocumentsHead)
        Me.pnlActivitiesDocuments.Controls.Add(Me.lblActivitiesDocumentsDisplay)
        Me.pnlActivitiesDocuments.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlActivitiesDocuments.Location = New System.Drawing.Point(0, 1064)
        Me.pnlActivitiesDocuments.Name = "pnlActivitiesDocuments"
        Me.pnlActivitiesDocuments.Size = New System.Drawing.Size(996, 24)
        Me.pnlActivitiesDocuments.TabIndex = 194
        '
        'lblActivitiesDocumentsHead
        '
        Me.lblActivitiesDocumentsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblActivitiesDocumentsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblActivitiesDocumentsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblActivitiesDocumentsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblActivitiesDocumentsHead.Name = "lblActivitiesDocumentsHead"
        Me.lblActivitiesDocumentsHead.Size = New System.Drawing.Size(980, 24)
        Me.lblActivitiesDocumentsHead.TabIndex = 1
        Me.lblActivitiesDocumentsHead.Text = "Activities / Documents"
        Me.lblActivitiesDocumentsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblActivitiesDocumentsDisplay
        '
        Me.lblActivitiesDocumentsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblActivitiesDocumentsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblActivitiesDocumentsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblActivitiesDocumentsDisplay.Name = "lblActivitiesDocumentsDisplay"
        Me.lblActivitiesDocumentsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblActivitiesDocumentsDisplay.TabIndex = 0
        Me.lblActivitiesDocumentsDisplay.Text = "-"
        Me.lblActivitiesDocumentsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlFundsEligibilityDetails
        '
        Me.pnlFundsEligibilityDetails.AutoScroll = True
        Me.pnlFundsEligibilityDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.pnlFundsEligibilityQuestions)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblEligibilityComments)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.txtEligibilityComments)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.btnSendtoPM)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.btnViewCheckList)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblQuestionsNA)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblQuestionsNo)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblQuestionsYes)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblQuestions)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.pnlTFAssess)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.pnlFundsEligibilitySysQuestions)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblSysQuestionsYes)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblSysQuestions)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblSysQuestionsNA)
        Me.pnlFundsEligibilityDetails.Controls.Add(Me.lblSysQuestionsNo)
        Me.pnlFundsEligibilityDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFundsEligibilityDetails.Location = New System.Drawing.Point(0, 632)
        Me.pnlFundsEligibilityDetails.Name = "pnlFundsEligibilityDetails"
        Me.pnlFundsEligibilityDetails.Size = New System.Drawing.Size(996, 432)
        Me.pnlFundsEligibilityDetails.TabIndex = 193
        '
        'pnlFundsEligibilityQuestions
        '
        Me.pnlFundsEligibilityQuestions.AutoScroll = True
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion17NA)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion17No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion17Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion17)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion16NA)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion16No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion16Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion16)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion15NA)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion15No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion15Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion15)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion14NA)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion14No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion14Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion14)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion10NA)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion10No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion10Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion10)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion8NA)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion8No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion8Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion8)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion6NA)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion6No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion6Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion6)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion5NA)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion5No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion5Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion2No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion2Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion1No)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.chkQuestion1Yes)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion5)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion2)
        Me.pnlFundsEligibilityQuestions.Controls.Add(Me.lblQuestion1)
        Me.pnlFundsEligibilityQuestions.Location = New System.Drawing.Point(24, 40)
        Me.pnlFundsEligibilityQuestions.Name = "pnlFundsEligibilityQuestions"
        Me.pnlFundsEligibilityQuestions.Size = New System.Drawing.Size(528, 248)
        Me.pnlFundsEligibilityQuestions.TabIndex = 0
        '
        'chkQuestion17NA
        '
        Me.chkQuestion17NA.Location = New System.Drawing.Point(480, 224)
        Me.chkQuestion17NA.Name = "chkQuestion17NA"
        Me.chkQuestion17NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion17NA.TabIndex = 27
        '
        'chkQuestion17No
        '
        Me.chkQuestion17No.Location = New System.Drawing.Point(448, 224)
        Me.chkQuestion17No.Name = "chkQuestion17No"
        Me.chkQuestion17No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion17No.TabIndex = 26
        '
        'chkQuestion17Yes
        '
        Me.chkQuestion17Yes.Location = New System.Drawing.Point(416, 224)
        Me.chkQuestion17Yes.Name = "chkQuestion17Yes"
        Me.chkQuestion17Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion17Yes.TabIndex = 25
        '
        'lblQuestion17
        '
        Me.lblQuestion17.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion17.Location = New System.Drawing.Point(16, 224)
        Me.lblQuestion17.Name = "lblQuestion17"
        Me.lblQuestion17.Size = New System.Drawing.Size(368, 16)
        Me.lblQuestion17.TabIndex = 247
        Me.lblQuestion17.Text = "17. Has the regional office or project manager inspected the site?"
        Me.lblQuestion17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion16NA
        '
        Me.chkQuestion16NA.Location = New System.Drawing.Point(480, 200)
        Me.chkQuestion16NA.Name = "chkQuestion16NA"
        Me.chkQuestion16NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion16NA.TabIndex = 24
        '
        'chkQuestion16No
        '
        Me.chkQuestion16No.Location = New System.Drawing.Point(448, 200)
        Me.chkQuestion16No.Name = "chkQuestion16No"
        Me.chkQuestion16No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion16No.TabIndex = 23
        '
        'chkQuestion16Yes
        '
        Me.chkQuestion16Yes.Location = New System.Drawing.Point(416, 200)
        Me.chkQuestion16Yes.Name = "chkQuestion16Yes"
        Me.chkQuestion16Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion16Yes.TabIndex = 22
        '
        'lblQuestion16
        '
        Me.lblQuestion16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion16.Location = New System.Drawing.Point(16, 200)
        Me.lblQuestion16.Name = "lblQuestion16"
        Me.lblQuestion16.Size = New System.Drawing.Size(328, 17)
        Me.lblQuestion16.TabIndex = 243
        Me.lblQuestion16.Text = "16. Has the owner performed the annual ALLD tests?"
        Me.lblQuestion16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion15NA
        '
        Me.chkQuestion15NA.Location = New System.Drawing.Point(480, 176)
        Me.chkQuestion15NA.Name = "chkQuestion15NA"
        Me.chkQuestion15NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion15NA.TabIndex = 21
        '
        'chkQuestion15No
        '
        Me.chkQuestion15No.Location = New System.Drawing.Point(448, 176)
        Me.chkQuestion15No.Name = "chkQuestion15No"
        Me.chkQuestion15No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion15No.TabIndex = 20
        '
        'chkQuestion15Yes
        '
        Me.chkQuestion15Yes.Location = New System.Drawing.Point(416, 176)
        Me.chkQuestion15Yes.Name = "chkQuestion15Yes"
        Me.chkQuestion15Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion15Yes.TabIndex = 19
        '
        'lblQuestion15
        '
        Me.lblQuestion15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion15.Location = New System.Drawing.Point(16, 176)
        Me.lblQuestion15.Name = "lblQuestion15"
        Me.lblQuestion15.Size = New System.Drawing.Size(336, 17)
        Me.lblQuestion15.TabIndex = 239
        Me.lblQuestion15.Text = "15. Has the owner performed a PTT on the tanks and lines?"
        Me.lblQuestion15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion14NA
        '
        Me.chkQuestion14NA.Location = New System.Drawing.Point(480, 152)
        Me.chkQuestion14NA.Name = "chkQuestion14NA"
        Me.chkQuestion14NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion14NA.TabIndex = 18
        '
        'chkQuestion14No
        '
        Me.chkQuestion14No.Location = New System.Drawing.Point(448, 152)
        Me.chkQuestion14No.Name = "chkQuestion14No"
        Me.chkQuestion14No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion14No.TabIndex = 17
        '
        'chkQuestion14Yes
        '
        Me.chkQuestion14Yes.Location = New System.Drawing.Point(416, 152)
        Me.chkQuestion14Yes.Name = "chkQuestion14Yes"
        Me.chkQuestion14Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion14Yes.TabIndex = 16
        '
        'lblQuestion14
        '
        Me.lblQuestion14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion14.Location = New System.Drawing.Point(16, 152)
        Me.lblQuestion14.Name = "lblQuestion14"
        Me.lblQuestion14.Size = New System.Drawing.Size(376, 17)
        Me.lblQuestion14.TabIndex = 235
        Me.lblQuestion14.Text = "14. Have the tanks complied with temporary closure requirements?"
        Me.lblQuestion14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion10NA
        '
        Me.chkQuestion10NA.Location = New System.Drawing.Point(480, 128)
        Me.chkQuestion10NA.Name = "chkQuestion10NA"
        Me.chkQuestion10NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion10NA.TabIndex = 15
        '
        'chkQuestion10No
        '
        Me.chkQuestion10No.Location = New System.Drawing.Point(448, 128)
        Me.chkQuestion10No.Name = "chkQuestion10No"
        Me.chkQuestion10No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion10No.TabIndex = 14
        '
        'chkQuestion10Yes
        '
        Me.chkQuestion10Yes.Location = New System.Drawing.Point(416, 128)
        Me.chkQuestion10Yes.Name = "chkQuestion10Yes"
        Me.chkQuestion10Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion10Yes.TabIndex = 13
        '
        'lblQuestion10
        '
        Me.lblQuestion10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion10.Location = New System.Drawing.Point(16, 128)
        Me.lblQuestion10.Name = "lblQuestion10"
        Me.lblQuestion10.Size = New System.Drawing.Size(352, 17)
        Me.lblQuestion10.TabIndex = 231
        Me.lblQuestion10.Text = "10. Have pipe leak detection records been submitted?"
        Me.lblQuestion10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion8NA
        '
        Me.chkQuestion8NA.Location = New System.Drawing.Point(480, 104)
        Me.chkQuestion8NA.Name = "chkQuestion8NA"
        Me.chkQuestion8NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion8NA.TabIndex = 12
        '
        'chkQuestion8No
        '
        Me.chkQuestion8No.Location = New System.Drawing.Point(448, 104)
        Me.chkQuestion8No.Name = "chkQuestion8No"
        Me.chkQuestion8No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion8No.TabIndex = 11
        '
        'chkQuestion8Yes
        '
        Me.chkQuestion8Yes.Location = New System.Drawing.Point(416, 104)
        Me.chkQuestion8Yes.Name = "chkQuestion8Yes"
        Me.chkQuestion8Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion8Yes.TabIndex = 10
        '
        'lblQuestion8
        '
        Me.lblQuestion8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion8.Location = New System.Drawing.Point(16, 104)
        Me.lblQuestion8.Name = "lblQuestion8"
        Me.lblQuestion8.Size = New System.Drawing.Size(344, 17)
        Me.lblQuestion8.TabIndex = 227
        Me.lblQuestion8.Text = "8. Have tank leak detection records been submitted?"
        Me.lblQuestion8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion6NA
        '
        Me.chkQuestion6NA.Location = New System.Drawing.Point(480, 80)
        Me.chkQuestion6NA.Name = "chkQuestion6NA"
        Me.chkQuestion6NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion6NA.TabIndex = 9
        '
        'chkQuestion6No
        '
        Me.chkQuestion6No.Location = New System.Drawing.Point(448, 80)
        Me.chkQuestion6No.Name = "chkQuestion6No"
        Me.chkQuestion6No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion6No.TabIndex = 8
        '
        'chkQuestion6Yes
        '
        Me.chkQuestion6Yes.Location = New System.Drawing.Point(416, 80)
        Me.chkQuestion6Yes.Name = "chkQuestion6Yes"
        Me.chkQuestion6Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion6Yes.TabIndex = 7
        '
        'lblQuestion6
        '
        Me.lblQuestion6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion6.Location = New System.Drawing.Point(16, 80)
        Me.lblQuestion6.Name = "lblQuestion6"
        Me.lblQuestion6.Size = New System.Drawing.Size(376, 17)
        Me.lblQuestion6.TabIndex = 223
        Me.lblQuestion6.Text = "6. Was the release reported verbally within 24 hours of discovery?"
        Me.lblQuestion6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion5NA
        '
        Me.chkQuestion5NA.Location = New System.Drawing.Point(480, 56)
        Me.chkQuestion5NA.Name = "chkQuestion5NA"
        Me.chkQuestion5NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion5NA.TabIndex = 6
        '
        'chkQuestion5No
        '
        Me.chkQuestion5No.Location = New System.Drawing.Point(448, 56)
        Me.chkQuestion5No.Name = "chkQuestion5No"
        Me.chkQuestion5No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion5No.TabIndex = 5
        '
        'chkQuestion5Yes
        '
        Me.chkQuestion5Yes.Location = New System.Drawing.Point(416, 56)
        Me.chkQuestion5Yes.Name = "chkQuestion5Yes"
        Me.chkQuestion5Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion5Yes.TabIndex = 4
        '
        'chkQuestion2No
        '
        Me.chkQuestion2No.Location = New System.Drawing.Point(448, 32)
        Me.chkQuestion2No.Name = "chkQuestion2No"
        Me.chkQuestion2No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion2No.TabIndex = 3
        '
        'chkQuestion2Yes
        '
        Me.chkQuestion2Yes.Location = New System.Drawing.Point(416, 32)
        Me.chkQuestion2Yes.Name = "chkQuestion2Yes"
        Me.chkQuestion2Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion2Yes.TabIndex = 2
        '
        'chkQuestion1No
        '
        Me.chkQuestion1No.Location = New System.Drawing.Point(448, 8)
        Me.chkQuestion1No.Name = "chkQuestion1No"
        Me.chkQuestion1No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion1No.TabIndex = 1
        '
        'chkQuestion1Yes
        '
        Me.chkQuestion1Yes.Location = New System.Drawing.Point(416, 8)
        Me.chkQuestion1Yes.Name = "chkQuestion1Yes"
        Me.chkQuestion1Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion1Yes.TabIndex = 0
        '
        'lblQuestion5
        '
        Me.lblQuestion5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion5.Location = New System.Drawing.Point(16, 56)
        Me.lblQuestion5.Name = "lblQuestion5"
        Me.lblQuestion5.Size = New System.Drawing.Size(352, 17)
        Me.lblQuestion5.TabIndex = 211
        Me.lblQuestion5.Text = "5. Has the owner submitted a written confirmation of a release?"
        Me.lblQuestion5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblQuestion2
        '
        Me.lblQuestion2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion2.Location = New System.Drawing.Point(16, 32)
        Me.lblQuestion2.Name = "lblQuestion2"
        Me.lblQuestion2.Size = New System.Drawing.Size(296, 17)
        Me.lblQuestion2.TabIndex = 210
        Me.lblQuestion2.Text = "2. Are all known USTs on site properly registered?"
        Me.lblQuestion2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblQuestion1
        '
        Me.lblQuestion1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion1.Location = New System.Drawing.Point(16, 8)
        Me.lblQuestion1.Name = "lblQuestion1"
        Me.lblQuestion1.Size = New System.Drawing.Size(272, 17)
        Me.lblQuestion1.TabIndex = 209
        Me.lblQuestion1.Text = "1. Has the owner submitted a notification form?"
        Me.lblQuestion1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEligibilityComments
        '
        Me.lblEligibilityComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEligibilityComments.Location = New System.Drawing.Point(600, 136)
        Me.lblEligibilityComments.Name = "lblEligibilityComments"
        Me.lblEligibilityComments.Size = New System.Drawing.Size(136, 17)
        Me.lblEligibilityComments.TabIndex = 214
        Me.lblEligibilityComments.Text = "Eligibility Comments"
        Me.lblEligibilityComments.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtEligibilityComments
        '
        Me.txtEligibilityComments.AcceptsTab = True
        Me.txtEligibilityComments.AutoSize = False
        Me.txtEligibilityComments.Location = New System.Drawing.Point(600, 160)
        Me.txtEligibilityComments.Multiline = True
        Me.txtEligibilityComments.Name = "txtEligibilityComments"
        Me.txtEligibilityComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEligibilityComments.Size = New System.Drawing.Size(232, 80)
        Me.txtEligibilityComments.TabIndex = 3
        Me.txtEligibilityComments.Text = ""
        '
        'btnSendtoPM
        '
        Me.btnSendtoPM.Location = New System.Drawing.Point(600, 80)
        Me.btnSendtoPM.Name = "btnSendtoPM"
        Me.btnSendtoPM.Size = New System.Drawing.Size(112, 40)
        Me.btnSendtoPM.TabIndex = 2
        Me.btnSendtoPM.Text = "Send to PM-Head for Review"
        '
        'btnViewCheckList
        '
        Me.btnViewCheckList.Location = New System.Drawing.Point(600, 32)
        Me.btnViewCheckList.Name = "btnViewCheckList"
        Me.btnViewCheckList.Size = New System.Drawing.Size(112, 40)
        Me.btnViewCheckList.TabIndex = 1
        Me.btnViewCheckList.Text = "View/Print Checklist"
        '
        'lblQuestionsNA
        '
        Me.lblQuestionsNA.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestionsNA.Location = New System.Drawing.Point(496, 16)
        Me.lblQuestionsNA.Name = "lblQuestionsNA"
        Me.lblQuestionsNA.Size = New System.Drawing.Size(24, 17)
        Me.lblQuestionsNA.TabIndex = 210
        Me.lblQuestionsNA.Text = "NA"
        Me.lblQuestionsNA.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblQuestionsNo
        '
        Me.lblQuestionsNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestionsNo.Location = New System.Drawing.Point(464, 16)
        Me.lblQuestionsNo.Name = "lblQuestionsNo"
        Me.lblQuestionsNo.Size = New System.Drawing.Size(24, 17)
        Me.lblQuestionsNo.TabIndex = 203
        Me.lblQuestionsNo.Text = "No"
        Me.lblQuestionsNo.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblQuestionsYes
        '
        Me.lblQuestionsYes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestionsYes.Location = New System.Drawing.Point(424, 16)
        Me.lblQuestionsYes.Name = "lblQuestionsYes"
        Me.lblQuestionsYes.Size = New System.Drawing.Size(32, 17)
        Me.lblQuestionsYes.TabIndex = 139
        Me.lblQuestionsYes.Text = "Yes"
        Me.lblQuestionsYes.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblQuestions
        '
        Me.lblQuestions.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestions.Location = New System.Drawing.Point(40, 16)
        Me.lblQuestions.Name = "lblQuestions"
        Me.lblQuestions.Size = New System.Drawing.Size(72, 17)
        Me.lblQuestions.TabIndex = 138
        Me.lblQuestions.Text = "Questions"
        Me.lblQuestions.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlTFAssess
        '
        Me.pnlTFAssess.Controls.Add(Me.dtCommissionOn)
        Me.pnlTFAssess.Controls.Add(Me.dtOPCHeadOn)
        Me.pnlTFAssess.Controls.Add(Me.dtUSTChiefOn)
        Me.pnlTFAssess.Controls.Add(Me.dtPMHeadOn)
        Me.pnlTFAssess.Controls.Add(Me.chkforCommission)
        Me.pnlTFAssess.Controls.Add(Me.chkforHeadofOPC)
        Me.pnlTFAssess.Controls.Add(Me.lblCommissionBy)
        Me.pnlTFAssess.Controls.Add(Me.lblOPCHeadBy)
        Me.pnlTFAssess.Controls.Add(Me.lblUSTChiefBy)
        Me.pnlTFAssess.Controls.Add(Me.txtCommissionBy)
        Me.pnlTFAssess.Controls.Add(Me.txtOPCHeadBy)
        Me.pnlTFAssess.Controls.Add(Me.txtUSTChiefBy)
        Me.pnlTFAssess.Controls.Add(Me.txtPMHeadBy)
        Me.pnlTFAssess.Controls.Add(Me.lblPMHeadBy)
        Me.pnlTFAssess.Controls.Add(Me.lblCommissionOn)
        Me.pnlTFAssess.Controls.Add(Me.lblOPCHeadOn)
        Me.pnlTFAssess.Controls.Add(Me.lblUSTChiefOn)
        Me.pnlTFAssess.Controls.Add(Me.lblPMHeadOn)
        Me.pnlTFAssess.Controls.Add(Me.chkCommissionNo)
        Me.pnlTFAssess.Controls.Add(Me.chkCommissionYes)
        Me.pnlTFAssess.Controls.Add(Me.lblCommission)
        Me.pnlTFAssess.Controls.Add(Me.chkOPCHeadNo)
        Me.pnlTFAssess.Controls.Add(Me.chkOPCHeadYes)
        Me.pnlTFAssess.Controls.Add(Me.lblOPCHead)
        Me.pnlTFAssess.Controls.Add(Me.chkUSTChiefUndecided)
        Me.pnlTFAssess.Controls.Add(Me.chkUSTChiefNo)
        Me.pnlTFAssess.Controls.Add(Me.chkUSTChiefYes)
        Me.pnlTFAssess.Controls.Add(Me.lblUSTChief)
        Me.pnlTFAssess.Controls.Add(Me.lblUndecided)
        Me.pnlTFAssess.Controls.Add(Me.lblNo)
        Me.pnlTFAssess.Controls.Add(Me.lblYes)
        Me.pnlTFAssess.Controls.Add(Me.chkPMHeadUndecided)
        Me.pnlTFAssess.Controls.Add(Me.chkPMHeadYes)
        Me.pnlTFAssess.Controls.Add(Me.lblPMHead)
        Me.pnlTFAssess.Location = New System.Drawing.Point(24, 296)
        Me.pnlTFAssess.Name = "pnlTFAssess"
        Me.pnlTFAssess.Size = New System.Drawing.Size(840, 128)
        Me.pnlTFAssess.TabIndex = 261
        '
        'dtCommissionOn
        '
        Me.dtCommissionOn.Checked = False
        Me.dtCommissionOn.CustomFormat = "__/__/____"
        Me.dtCommissionOn.Enabled = False
        Me.dtCommissionOn.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtCommissionOn.Location = New System.Drawing.Point(376, 102)
        Me.dtCommissionOn.Name = "dtCommissionOn"
        Me.dtCommissionOn.ShowCheckBox = True
        Me.dtCommissionOn.Size = New System.Drawing.Size(104, 21)
        Me.dtCommissionOn.TabIndex = 293
        '
        'dtOPCHeadOn
        '
        Me.dtOPCHeadOn.Checked = False
        Me.dtOPCHeadOn.CustomFormat = "__/__/____"
        Me.dtOPCHeadOn.Enabled = False
        Me.dtOPCHeadOn.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtOPCHeadOn.Location = New System.Drawing.Point(376, 78)
        Me.dtOPCHeadOn.Name = "dtOPCHeadOn"
        Me.dtOPCHeadOn.ShowCheckBox = True
        Me.dtOPCHeadOn.Size = New System.Drawing.Size(104, 21)
        Me.dtOPCHeadOn.TabIndex = 292
        '
        'dtUSTChiefOn
        '
        Me.dtUSTChiefOn.Checked = False
        Me.dtUSTChiefOn.CustomFormat = "__/__/____"
        Me.dtUSTChiefOn.Enabled = False
        Me.dtUSTChiefOn.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtUSTChiefOn.Location = New System.Drawing.Point(376, 54)
        Me.dtUSTChiefOn.Name = "dtUSTChiefOn"
        Me.dtUSTChiefOn.ShowCheckBox = True
        Me.dtUSTChiefOn.Size = New System.Drawing.Size(104, 21)
        Me.dtUSTChiefOn.TabIndex = 291
        '
        'dtPMHeadOn
        '
        Me.dtPMHeadOn.Checked = False
        Me.dtPMHeadOn.CustomFormat = "__/__/____"
        Me.dtPMHeadOn.Enabled = False
        Me.dtPMHeadOn.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPMHeadOn.Location = New System.Drawing.Point(376, 30)
        Me.dtPMHeadOn.Name = "dtPMHeadOn"
        Me.dtPMHeadOn.ShowCheckBox = True
        Me.dtPMHeadOn.Size = New System.Drawing.Size(104, 21)
        Me.dtPMHeadOn.TabIndex = 290
        '
        'chkforCommission
        '
        Me.chkforCommission.Location = New System.Drawing.Point(632, 102)
        Me.chkforCommission.Name = "chkforCommission"
        Me.chkforCommission.Size = New System.Drawing.Size(120, 16)
        Me.chkforCommission.TabIndex = 274
        Me.chkforCommission.Text = "for Commission"
        '
        'chkforHeadofOPC
        '
        Me.chkforHeadofOPC.Location = New System.Drawing.Point(632, 78)
        Me.chkforHeadofOPC.Name = "chkforHeadofOPC"
        Me.chkforHeadofOPC.Size = New System.Drawing.Size(120, 16)
        Me.chkforHeadofOPC.TabIndex = 270
        Me.chkforHeadofOPC.Text = "for Head of OPC"
        '
        'lblCommissionBy
        '
        Me.lblCommissionBy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCommissionBy.Location = New System.Drawing.Point(488, 102)
        Me.lblCommissionBy.Name = "lblCommissionBy"
        Me.lblCommissionBy.Size = New System.Drawing.Size(32, 16)
        Me.lblCommissionBy.TabIndex = 289
        Me.lblCommissionBy.Text = "By:"
        Me.lblCommissionBy.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblOPCHeadBy
        '
        Me.lblOPCHeadBy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOPCHeadBy.Location = New System.Drawing.Point(488, 78)
        Me.lblOPCHeadBy.Name = "lblOPCHeadBy"
        Me.lblOPCHeadBy.Size = New System.Drawing.Size(32, 16)
        Me.lblOPCHeadBy.TabIndex = 288
        Me.lblOPCHeadBy.Text = "By:"
        Me.lblOPCHeadBy.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblUSTChiefBy
        '
        Me.lblUSTChiefBy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUSTChiefBy.Location = New System.Drawing.Point(488, 54)
        Me.lblUSTChiefBy.Name = "lblUSTChiefBy"
        Me.lblUSTChiefBy.Size = New System.Drawing.Size(32, 16)
        Me.lblUSTChiefBy.TabIndex = 287
        Me.lblUSTChiefBy.Text = "By:"
        Me.lblUSTChiefBy.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtCommissionBy
        '
        Me.txtCommissionBy.AcceptsTab = True
        Me.txtCommissionBy.AutoSize = False
        Me.txtCommissionBy.Location = New System.Drawing.Point(520, 102)
        Me.txtCommissionBy.Name = "txtCommissionBy"
        Me.txtCommissionBy.ReadOnly = True
        Me.txtCommissionBy.Size = New System.Drawing.Size(96, 21)
        Me.txtCommissionBy.TabIndex = 273
        Me.txtCommissionBy.Text = ""
        Me.txtCommissionBy.WordWrap = False
        '
        'txtOPCHeadBy
        '
        Me.txtOPCHeadBy.AcceptsTab = True
        Me.txtOPCHeadBy.AutoSize = False
        Me.txtOPCHeadBy.Location = New System.Drawing.Point(520, 78)
        Me.txtOPCHeadBy.Name = "txtOPCHeadBy"
        Me.txtOPCHeadBy.ReadOnly = True
        Me.txtOPCHeadBy.Size = New System.Drawing.Size(96, 21)
        Me.txtOPCHeadBy.TabIndex = 269
        Me.txtOPCHeadBy.Text = ""
        Me.txtOPCHeadBy.WordWrap = False
        '
        'txtUSTChiefBy
        '
        Me.txtUSTChiefBy.AcceptsTab = True
        Me.txtUSTChiefBy.AutoSize = False
        Me.txtUSTChiefBy.Location = New System.Drawing.Point(520, 54)
        Me.txtUSTChiefBy.Name = "txtUSTChiefBy"
        Me.txtUSTChiefBy.ReadOnly = True
        Me.txtUSTChiefBy.Size = New System.Drawing.Size(96, 21)
        Me.txtUSTChiefBy.TabIndex = 266
        Me.txtUSTChiefBy.Text = ""
        Me.txtUSTChiefBy.WordWrap = False
        '
        'txtPMHeadBy
        '
        Me.txtPMHeadBy.AcceptsTab = True
        Me.txtPMHeadBy.AutoSize = False
        Me.txtPMHeadBy.Location = New System.Drawing.Point(520, 30)
        Me.txtPMHeadBy.Name = "txtPMHeadBy"
        Me.txtPMHeadBy.ReadOnly = True
        Me.txtPMHeadBy.Size = New System.Drawing.Size(96, 21)
        Me.txtPMHeadBy.TabIndex = 262
        Me.txtPMHeadBy.Text = ""
        Me.txtPMHeadBy.WordWrap = False
        '
        'lblPMHeadBy
        '
        Me.lblPMHeadBy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPMHeadBy.Location = New System.Drawing.Point(488, 30)
        Me.lblPMHeadBy.Name = "lblPMHeadBy"
        Me.lblPMHeadBy.Size = New System.Drawing.Size(32, 16)
        Me.lblPMHeadBy.TabIndex = 286
        Me.lblPMHeadBy.Text = "By:"
        Me.lblPMHeadBy.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblCommissionOn
        '
        Me.lblCommissionOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCommissionOn.Location = New System.Drawing.Point(344, 102)
        Me.lblCommissionOn.Name = "lblCommissionOn"
        Me.lblCommissionOn.Size = New System.Drawing.Size(32, 16)
        Me.lblCommissionOn.TabIndex = 285
        Me.lblCommissionOn.Text = "On:"
        Me.lblCommissionOn.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblOPCHeadOn
        '
        Me.lblOPCHeadOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOPCHeadOn.Location = New System.Drawing.Point(344, 78)
        Me.lblOPCHeadOn.Name = "lblOPCHeadOn"
        Me.lblOPCHeadOn.Size = New System.Drawing.Size(32, 16)
        Me.lblOPCHeadOn.TabIndex = 284
        Me.lblOPCHeadOn.Text = "On:"
        Me.lblOPCHeadOn.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblUSTChiefOn
        '
        Me.lblUSTChiefOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUSTChiefOn.Location = New System.Drawing.Point(344, 54)
        Me.lblUSTChiefOn.Name = "lblUSTChiefOn"
        Me.lblUSTChiefOn.Size = New System.Drawing.Size(32, 16)
        Me.lblUSTChiefOn.TabIndex = 283
        Me.lblUSTChiefOn.Text = "On:"
        Me.lblUSTChiefOn.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblPMHeadOn
        '
        Me.lblPMHeadOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPMHeadOn.Location = New System.Drawing.Point(344, 30)
        Me.lblPMHeadOn.Name = "lblPMHeadOn"
        Me.lblPMHeadOn.Size = New System.Drawing.Size(32, 16)
        Me.lblPMHeadOn.TabIndex = 282
        Me.lblPMHeadOn.Text = "On:"
        Me.lblPMHeadOn.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'chkCommissionNo
        '
        Me.chkCommissionNo.Location = New System.Drawing.Point(248, 102)
        Me.chkCommissionNo.Name = "chkCommissionNo"
        Me.chkCommissionNo.Size = New System.Drawing.Size(16, 16)
        Me.chkCommissionNo.TabIndex = 272
        '
        'chkCommissionYes
        '
        Me.chkCommissionYes.Location = New System.Drawing.Point(216, 102)
        Me.chkCommissionYes.Name = "chkCommissionYes"
        Me.chkCommissionYes.Size = New System.Drawing.Size(16, 16)
        Me.chkCommissionYes.TabIndex = 271
        '
        'lblCommission
        '
        Me.lblCommission.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCommission.Location = New System.Drawing.Point(8, 102)
        Me.lblCommission.Name = "lblCommission"
        Me.lblCommission.Size = New System.Drawing.Size(192, 17)
        Me.lblCommission.TabIndex = 281
        Me.lblCommission.Text = "Commission - Is the site eligible?"
        Me.lblCommission.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkOPCHeadNo
        '
        Me.chkOPCHeadNo.Location = New System.Drawing.Point(248, 78)
        Me.chkOPCHeadNo.Name = "chkOPCHeadNo"
        Me.chkOPCHeadNo.Size = New System.Drawing.Size(16, 16)
        Me.chkOPCHeadNo.TabIndex = 268
        '
        'chkOPCHeadYes
        '
        Me.chkOPCHeadYes.Location = New System.Drawing.Point(216, 78)
        Me.chkOPCHeadYes.Name = "chkOPCHeadYes"
        Me.chkOPCHeadYes.Size = New System.Drawing.Size(16, 16)
        Me.chkOPCHeadYes.TabIndex = 267
        '
        'lblOPCHead
        '
        Me.lblOPCHead.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOPCHead.Location = New System.Drawing.Point(8, 78)
        Me.lblOPCHead.Name = "lblOPCHead"
        Me.lblOPCHead.Size = New System.Drawing.Size(184, 17)
        Me.lblOPCHead.TabIndex = 280
        Me.lblOPCHead.Text = "OPC-Head - Is the site eligible?"
        Me.lblOPCHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkUSTChiefUndecided
        '
        Me.chkUSTChiefUndecided.Location = New System.Drawing.Point(280, 54)
        Me.chkUSTChiefUndecided.Name = "chkUSTChiefUndecided"
        Me.chkUSTChiefUndecided.Size = New System.Drawing.Size(16, 16)
        Me.chkUSTChiefUndecided.TabIndex = 265
        '
        'chkUSTChiefNo
        '
        Me.chkUSTChiefNo.Location = New System.Drawing.Point(248, 54)
        Me.chkUSTChiefNo.Name = "chkUSTChiefNo"
        Me.chkUSTChiefNo.Size = New System.Drawing.Size(16, 16)
        Me.chkUSTChiefNo.TabIndex = 264
        '
        'chkUSTChiefYes
        '
        Me.chkUSTChiefYes.Location = New System.Drawing.Point(216, 54)
        Me.chkUSTChiefYes.Name = "chkUSTChiefYes"
        Me.chkUSTChiefYes.Size = New System.Drawing.Size(16, 16)
        Me.chkUSTChiefYes.TabIndex = 263
        '
        'lblUSTChief
        '
        Me.lblUSTChief.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUSTChief.Location = New System.Drawing.Point(8, 54)
        Me.lblUSTChief.Name = "lblUSTChief"
        Me.lblUSTChief.Size = New System.Drawing.Size(176, 17)
        Me.lblUSTChief.TabIndex = 279
        Me.lblUSTChief.Text = "UST Chief - Is the site eligible?"
        Me.lblUSTChief.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblUndecided
        '
        Me.lblUndecided.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUndecided.Location = New System.Drawing.Point(272, 6)
        Me.lblUndecided.Name = "lblUndecided"
        Me.lblUndecided.Size = New System.Drawing.Size(96, 17)
        Me.lblUndecided.TabIndex = 278
        Me.lblUndecided.Text = "Undecided"
        Me.lblUndecided.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblNo
        '
        Me.lblNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNo.Location = New System.Drawing.Point(240, 6)
        Me.lblNo.Name = "lblNo"
        Me.lblNo.Size = New System.Drawing.Size(24, 17)
        Me.lblNo.TabIndex = 277
        Me.lblNo.Text = "No"
        Me.lblNo.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblYes
        '
        Me.lblYes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYes.Location = New System.Drawing.Point(200, 6)
        Me.lblYes.Name = "lblYes"
        Me.lblYes.Size = New System.Drawing.Size(32, 17)
        Me.lblYes.TabIndex = 276
        Me.lblYes.Text = "Yes"
        Me.lblYes.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkPMHeadUndecided
        '
        Me.chkPMHeadUndecided.Location = New System.Drawing.Point(280, 30)
        Me.chkPMHeadUndecided.Name = "chkPMHeadUndecided"
        Me.chkPMHeadUndecided.Size = New System.Drawing.Size(16, 16)
        Me.chkPMHeadUndecided.TabIndex = 261
        '
        'chkPMHeadYes
        '
        Me.chkPMHeadYes.Location = New System.Drawing.Point(216, 30)
        Me.chkPMHeadYes.Name = "chkPMHeadYes"
        Me.chkPMHeadYes.Size = New System.Drawing.Size(16, 16)
        Me.chkPMHeadYes.TabIndex = 260
        '
        'lblPMHead
        '
        Me.lblPMHead.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPMHead.Location = New System.Drawing.Point(8, 30)
        Me.lblPMHead.Name = "lblPMHead"
        Me.lblPMHead.Size = New System.Drawing.Size(176, 17)
        Me.lblPMHead.TabIndex = 275
        Me.lblPMHead.Text = "PM-Head - Is the site eligible?"
        Me.lblPMHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlFundsEligibilitySysQuestions
        '
        Me.pnlFundsEligibilitySysQuestions.AutoScroll = True
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion13NA)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion13No)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion13Yes)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.lblQuestion13)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion12NA)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion12No)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion12Yes)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.lblQuestion12)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion11NA)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion11No)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion11Yes)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.lblQuestion11)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion9NA)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion9No)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion9Yes)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.lblQuestion9)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion7NA)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion7No)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion7Yes)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion4No)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion4Yes)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion3No)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion3Yes)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.lblQuestion7)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.lblQuestion4)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.lblQuestion3)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion4NA)
        Me.pnlFundsEligibilitySysQuestions.Controls.Add(Me.chkQuestion3NA)
        Me.pnlFundsEligibilitySysQuestions.Enabled = False
        Me.pnlFundsEligibilitySysQuestions.Location = New System.Drawing.Point(24, 456)
        Me.pnlFundsEligibilitySysQuestions.Name = "pnlFundsEligibilitySysQuestions"
        Me.pnlFundsEligibilitySysQuestions.Size = New System.Drawing.Size(552, 216)
        Me.pnlFundsEligibilitySysQuestions.TabIndex = 0
        '
        'chkQuestion13NA
        '
        Me.chkQuestion13NA.Enabled = False
        Me.chkQuestion13NA.Location = New System.Drawing.Point(504, 184)
        Me.chkQuestion13NA.Name = "chkQuestion13NA"
        Me.chkQuestion13NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion13NA.TabIndex = 18
        '
        'chkQuestion13No
        '
        Me.chkQuestion13No.Enabled = False
        Me.chkQuestion13No.Location = New System.Drawing.Point(472, 184)
        Me.chkQuestion13No.Name = "chkQuestion13No"
        Me.chkQuestion13No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion13No.TabIndex = 17
        '
        'chkQuestion13Yes
        '
        Me.chkQuestion13Yes.Enabled = False
        Me.chkQuestion13Yes.Location = New System.Drawing.Point(440, 184)
        Me.chkQuestion13Yes.Name = "chkQuestion13Yes"
        Me.chkQuestion13Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion13Yes.TabIndex = 16
        '
        'lblQuestion13
        '
        Me.lblQuestion13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion13.Location = New System.Drawing.Point(16, 184)
        Me.lblQuestion13.Name = "lblQuestion13"
        Me.lblQuestion13.Size = New System.Drawing.Size(352, 17)
        Me.lblQuestion13.TabIndex = 235
        Me.lblQuestion13.Text = "13. Does the UST system have the required overfill prevention?"
        Me.lblQuestion13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion12NA
        '
        Me.chkQuestion12NA.Enabled = False
        Me.chkQuestion12NA.Location = New System.Drawing.Point(504, 160)
        Me.chkQuestion12NA.Name = "chkQuestion12NA"
        Me.chkQuestion12NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion12NA.TabIndex = 15
        '
        'chkQuestion12No
        '
        Me.chkQuestion12No.Enabled = False
        Me.chkQuestion12No.Location = New System.Drawing.Point(472, 160)
        Me.chkQuestion12No.Name = "chkQuestion12No"
        Me.chkQuestion12No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion12No.TabIndex = 14
        '
        'chkQuestion12Yes
        '
        Me.chkQuestion12Yes.Enabled = False
        Me.chkQuestion12Yes.Location = New System.Drawing.Point(440, 160)
        Me.chkQuestion12Yes.Name = "chkQuestion12Yes"
        Me.chkQuestion12Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion12Yes.TabIndex = 13
        '
        'lblQuestion12
        '
        Me.lblQuestion12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion12.Location = New System.Drawing.Point(16, 160)
        Me.lblQuestion12.Name = "lblQuestion12"
        Me.lblQuestion12.Size = New System.Drawing.Size(344, 17)
        Me.lblQuestion12.TabIndex = 231
        Me.lblQuestion12.Text = "12. Does the UST system have the required spill prevention?"
        Me.lblQuestion12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion11NA
        '
        Me.chkQuestion11NA.Enabled = False
        Me.chkQuestion11NA.Location = New System.Drawing.Point(504, 136)
        Me.chkQuestion11NA.Name = "chkQuestion11NA"
        Me.chkQuestion11NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion11NA.TabIndex = 12
        '
        'chkQuestion11No
        '
        Me.chkQuestion11No.Enabled = False
        Me.chkQuestion11No.Location = New System.Drawing.Point(472, 136)
        Me.chkQuestion11No.Name = "chkQuestion11No"
        Me.chkQuestion11No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion11No.TabIndex = 11
        '
        'chkQuestion11Yes
        '
        Me.chkQuestion11Yes.Enabled = False
        Me.chkQuestion11Yes.Location = New System.Drawing.Point(440, 136)
        Me.chkQuestion11Yes.Name = "chkQuestion11Yes"
        Me.chkQuestion11Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion11Yes.TabIndex = 10
        '
        'lblQuestion11
        '
        Me.lblQuestion11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion11.Location = New System.Drawing.Point(16, 136)
        Me.lblQuestion11.Name = "lblQuestion11"
        Me.lblQuestion11.Size = New System.Drawing.Size(296, 17)
        Me.lblQuestion11.TabIndex = 227
        Me.lblQuestion11.Text = "11. Do the tanks and lines have corrosion protection?"
        Me.lblQuestion11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion9NA
        '
        Me.chkQuestion9NA.Enabled = False
        Me.chkQuestion9NA.Location = New System.Drawing.Point(504, 112)
        Me.chkQuestion9NA.Name = "chkQuestion9NA"
        Me.chkQuestion9NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion9NA.TabIndex = 9
        '
        'chkQuestion9No
        '
        Me.chkQuestion9No.Enabled = False
        Me.chkQuestion9No.Location = New System.Drawing.Point(472, 112)
        Me.chkQuestion9No.Name = "chkQuestion9No"
        Me.chkQuestion9No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion9No.TabIndex = 8
        '
        'chkQuestion9Yes
        '
        Me.chkQuestion9Yes.Enabled = False
        Me.chkQuestion9Yes.Location = New System.Drawing.Point(440, 112)
        Me.chkQuestion9Yes.Name = "chkQuestion9Yes"
        Me.chkQuestion9Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion9Yes.TabIndex = 7
        '
        'lblQuestion9
        '
        Me.lblQuestion9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion9.Location = New System.Drawing.Point(16, 96)
        Me.lblQuestion9.Name = "lblQuestion9"
        Me.lblQuestion9.Size = New System.Drawing.Size(312, 40)
        Me.lblQuestion9.TabIndex = 223
        Me.lblQuestion9.Text = "9. At the time of the release, was the site in compliance with the Compliance and" & _
        " Enforcement Division?"
        Me.lblQuestion9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion7NA
        '
        Me.chkQuestion7NA.Enabled = False
        Me.chkQuestion7NA.Location = New System.Drawing.Point(504, 72)
        Me.chkQuestion7NA.Name = "chkQuestion7NA"
        Me.chkQuestion7NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion7NA.TabIndex = 6
        '
        'chkQuestion7No
        '
        Me.chkQuestion7No.Enabled = False
        Me.chkQuestion7No.Location = New System.Drawing.Point(472, 72)
        Me.chkQuestion7No.Name = "chkQuestion7No"
        Me.chkQuestion7No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion7No.TabIndex = 5
        '
        'chkQuestion7Yes
        '
        Me.chkQuestion7Yes.Enabled = False
        Me.chkQuestion7Yes.Location = New System.Drawing.Point(440, 72)
        Me.chkQuestion7Yes.Name = "chkQuestion7Yes"
        Me.chkQuestion7Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion7Yes.TabIndex = 4
        '
        'chkQuestion4No
        '
        Me.chkQuestion4No.Enabled = False
        Me.chkQuestion4No.Location = New System.Drawing.Point(472, 48)
        Me.chkQuestion4No.Name = "chkQuestion4No"
        Me.chkQuestion4No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion4No.TabIndex = 3
        '
        'chkQuestion4Yes
        '
        Me.chkQuestion4Yes.Enabled = False
        Me.chkQuestion4Yes.Location = New System.Drawing.Point(440, 48)
        Me.chkQuestion4Yes.Name = "chkQuestion4Yes"
        Me.chkQuestion4Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion4Yes.TabIndex = 2
        '
        'chkQuestion3No
        '
        Me.chkQuestion3No.Enabled = False
        Me.chkQuestion3No.Location = New System.Drawing.Point(472, 24)
        Me.chkQuestion3No.Name = "chkQuestion3No"
        Me.chkQuestion3No.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion3No.TabIndex = 1
        '
        'chkQuestion3Yes
        '
        Me.chkQuestion3Yes.Enabled = False
        Me.chkQuestion3Yes.Location = New System.Drawing.Point(440, 24)
        Me.chkQuestion3Yes.Name = "chkQuestion3Yes"
        Me.chkQuestion3Yes.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion3Yes.TabIndex = 0
        '
        'lblQuestion7
        '
        Me.lblQuestion7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion7.Location = New System.Drawing.Point(16, 72)
        Me.lblQuestion7.Name = "lblQuestion7"
        Me.lblQuestion7.Size = New System.Drawing.Size(288, 17)
        Me.lblQuestion7.TabIndex = 211
        Me.lblQuestion7.Text = "7. Was the released product aviation or motor fuel?"
        Me.lblQuestion7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblQuestion4
        '
        Me.lblQuestion4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion4.Location = New System.Drawing.Point(16, 48)
        Me.lblQuestion4.Name = "lblQuestion4"
        Me.lblQuestion4.Size = New System.Drawing.Size(384, 17)
        Me.lblQuestion4.TabIndex = 210
        Me.lblQuestion4.Text = "4. Were tank fees (including late fees) paid prior to the release date?"
        Me.lblQuestion4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblQuestion3
        '
        Me.lblQuestion3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQuestion3.Location = New System.Drawing.Point(16, 8)
        Me.lblQuestion3.Name = "lblQuestion3"
        Me.lblQuestion3.Size = New System.Drawing.Size(416, 40)
        Me.lblQuestion3.TabIndex = 209
        Me.lblQuestion3.Text = "3. Has the site been active (UST owner can be identified and UST was/is in use fo" & _
        "r management and handling of motor fuel) on or after July 1, 1988?"
        Me.lblQuestion3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkQuestion4NA
        '
        Me.chkQuestion4NA.Enabled = False
        Me.chkQuestion4NA.Location = New System.Drawing.Point(504, 48)
        Me.chkQuestion4NA.Name = "chkQuestion4NA"
        Me.chkQuestion4NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion4NA.TabIndex = 6
        '
        'chkQuestion3NA
        '
        Me.chkQuestion3NA.Enabled = False
        Me.chkQuestion3NA.Location = New System.Drawing.Point(504, 24)
        Me.chkQuestion3NA.Name = "chkQuestion3NA"
        Me.chkQuestion3NA.Size = New System.Drawing.Size(16, 16)
        Me.chkQuestion3NA.TabIndex = 6
        '
        'lblSysQuestionsYes
        '
        Me.lblSysQuestionsYes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSysQuestionsYes.Location = New System.Drawing.Point(448, 432)
        Me.lblSysQuestionsYes.Name = "lblSysQuestionsYes"
        Me.lblSysQuestionsYes.Size = New System.Drawing.Size(32, 17)
        Me.lblSysQuestionsYes.TabIndex = 139
        Me.lblSysQuestionsYes.Text = "Yes"
        Me.lblSysQuestionsYes.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblSysQuestions
        '
        Me.lblSysQuestions.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSysQuestions.Location = New System.Drawing.Point(48, 432)
        Me.lblSysQuestions.Name = "lblSysQuestions"
        Me.lblSysQuestions.Size = New System.Drawing.Size(112, 17)
        Me.lblSysQuestions.TabIndex = 138
        Me.lblSysQuestions.Text = "System Questions"
        Me.lblSysQuestions.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSysQuestionsNA
        '
        Me.lblSysQuestionsNA.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSysQuestionsNA.Location = New System.Drawing.Point(520, 432)
        Me.lblSysQuestionsNA.Name = "lblSysQuestionsNA"
        Me.lblSysQuestionsNA.Size = New System.Drawing.Size(24, 17)
        Me.lblSysQuestionsNA.TabIndex = 210
        Me.lblSysQuestionsNA.Text = "NA"
        Me.lblSysQuestionsNA.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblSysQuestionsNo
        '
        Me.lblSysQuestionsNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSysQuestionsNo.Location = New System.Drawing.Point(488, 432)
        Me.lblSysQuestionsNo.Name = "lblSysQuestionsNo"
        Me.lblSysQuestionsNo.Size = New System.Drawing.Size(24, 17)
        Me.lblSysQuestionsNo.TabIndex = 203
        Me.lblSysQuestionsNo.Text = "No"
        Me.lblSysQuestionsNo.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'pnlFundsEligibility
        '
        Me.pnlFundsEligibility.Controls.Add(Me.lblFundsEligibilityHead)
        Me.pnlFundsEligibility.Controls.Add(Me.lblFundsEligibilityDisplay)
        Me.pnlFundsEligibility.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFundsEligibility.Location = New System.Drawing.Point(0, 608)
        Me.pnlFundsEligibility.Name = "pnlFundsEligibility"
        Me.pnlFundsEligibility.Size = New System.Drawing.Size(996, 24)
        Me.pnlFundsEligibility.TabIndex = 192
        '
        'lblFundsEligibilityHead
        '
        Me.lblFundsEligibilityHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblFundsEligibilityHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblFundsEligibilityHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblFundsEligibilityHead.Location = New System.Drawing.Point(16, 0)
        Me.lblFundsEligibilityHead.Name = "lblFundsEligibilityHead"
        Me.lblFundsEligibilityHead.Size = New System.Drawing.Size(980, 24)
        Me.lblFundsEligibilityHead.TabIndex = 1
        Me.lblFundsEligibilityHead.Text = "MGPTF Checklist"
        Me.lblFundsEligibilityHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFundsEligibilityDisplay
        '
        Me.lblFundsEligibilityDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFundsEligibilityDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblFundsEligibilityDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblFundsEligibilityDisplay.Name = "lblFundsEligibilityDisplay"
        Me.lblFundsEligibilityDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblFundsEligibilityDisplay.TabIndex = 0
        Me.lblFundsEligibilityDisplay.Text = "-"
        Me.lblFundsEligibilityDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PnlReleaseInfoDetails
        '
        Me.PnlReleaseInfoDetails.AutoScroll = True
        Me.PnlReleaseInfoDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PnlReleaseInfoDetails.Controls.Add(Me.btnEvtTankCollapse)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.btnEvtTankToggle)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.dtConfirmedOn)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.GroupBox2)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.GroupBox1)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.ugTankandPipes)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.cmbExtent)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.lblExtent)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.cmbLocation)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.cmbIdentifiedBy)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.lblLocation)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.lblIdentifiedBy)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.lblConfirmedOn)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.cmbCause)
        Me.PnlReleaseInfoDetails.Controls.Add(Me.lblCause)
        Me.PnlReleaseInfoDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.PnlReleaseInfoDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlReleaseInfoDetails.Location = New System.Drawing.Point(0, 176)
        Me.PnlReleaseInfoDetails.Name = "PnlReleaseInfoDetails"
        Me.PnlReleaseInfoDetails.Size = New System.Drawing.Size(996, 432)
        Me.PnlReleaseInfoDetails.TabIndex = 12
        '
        'btnEvtTankCollapse
        '
        Me.btnEvtTankCollapse.Location = New System.Drawing.Point(832, 224)
        Me.btnEvtTankCollapse.Name = "btnEvtTankCollapse"
        Me.btnEvtTankCollapse.Size = New System.Drawing.Size(112, 23)
        Me.btnEvtTankCollapse.TabIndex = 214
        Me.btnEvtTankCollapse.Text = "Collapse"
        '
        'btnEvtTankToggle
        '
        Me.btnEvtTankToggle.Location = New System.Drawing.Point(712, 224)
        Me.btnEvtTankToggle.Name = "btnEvtTankToggle"
        Me.btnEvtTankToggle.Size = New System.Drawing.Size(112, 23)
        Me.btnEvtTankToggle.TabIndex = 213
        Me.btnEvtTankToggle.Text = "Show Event Only"
        '
        'dtConfirmedOn
        '
        Me.dtConfirmedOn.Checked = False
        Me.dtConfirmedOn.CustomFormat = ""
        Me.dtConfirmedOn.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtConfirmedOn.Location = New System.Drawing.Point(96, 8)
        Me.dtConfirmedOn.Name = "dtConfirmedOn"
        Me.dtConfirmedOn.ShowCheckBox = True
        Me.dtConfirmedOn.Size = New System.Drawing.Size(104, 21)
        Me.dtConfirmedOn.TabIndex = 210
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkFreeProductUnKnown)
        Me.GroupBox2.Controls.Add(Me.chkFreeProductWasteOil)
        Me.GroupBox2.Controls.Add(Me.chkVaporPAH)
        Me.GroupBox2.Controls.Add(Me.chkVaporBTEX)
        Me.GroupBox2.Controls.Add(Me.chkToCVapor)
        Me.GroupBox2.Controls.Add(Me.chkFreeProductKerosene)
        Me.GroupBox2.Controls.Add(Me.chkFreeProductDiesel)
        Me.GroupBox2.Controls.Add(Me.chkFreeProductGasoline)
        Me.GroupBox2.Controls.Add(Me.chkToCFreeProduct)
        Me.GroupBox2.Controls.Add(Me.chkGroundWaterTPH)
        Me.GroupBox2.Controls.Add(Me.chkGroundWaterPAH)
        Me.GroupBox2.Controls.Add(Me.chkGroundWaterBTEX)
        Me.GroupBox2.Controls.Add(Me.chkGroundWater)
        Me.GroupBox2.Controls.Add(Me.chkSoilTPH)
        Me.GroupBox2.Controls.Add(Me.chkSoilPAH)
        Me.GroupBox2.Controls.Add(Me.chkSoilBTEX)
        Me.GroupBox2.Controls.Add(Me.chkSoil)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(384, 48)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(400, 168)
        Me.GroupBox2.TabIndex = 205
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Type of Contamination"
        '
        'chkFreeProductUnKnown
        '
        Me.chkFreeProductUnKnown.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFreeProductUnKnown.Location = New System.Drawing.Point(216, 144)
        Me.chkFreeProductUnKnown.Name = "chkFreeProductUnKnown"
        Me.chkFreeProductUnKnown.Size = New System.Drawing.Size(80, 16)
        Me.chkFreeProductUnKnown.TabIndex = 45
        Me.chkFreeProductUnKnown.Text = "UnKnown"
        '
        'chkFreeProductWasteOil
        '
        Me.chkFreeProductWasteOil.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFreeProductWasteOil.Location = New System.Drawing.Point(216, 120)
        Me.chkFreeProductWasteOil.Name = "chkFreeProductWasteOil"
        Me.chkFreeProductWasteOil.Size = New System.Drawing.Size(88, 16)
        Me.chkFreeProductWasteOil.TabIndex = 44
        Me.chkFreeProductWasteOil.Text = "Waste Oil"
        '
        'chkVaporPAH
        '
        Me.chkVaporPAH.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVaporPAH.Location = New System.Drawing.Point(328, 72)
        Me.chkVaporPAH.Name = "chkVaporPAH"
        Me.chkVaporPAH.Size = New System.Drawing.Size(56, 16)
        Me.chkVaporPAH.TabIndex = 48
        Me.chkVaporPAH.Text = "PAH"
        '
        'chkVaporBTEX
        '
        Me.chkVaporBTEX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVaporBTEX.Location = New System.Drawing.Point(328, 48)
        Me.chkVaporBTEX.Name = "chkVaporBTEX"
        Me.chkVaporBTEX.Size = New System.Drawing.Size(56, 16)
        Me.chkVaporBTEX.TabIndex = 47
        Me.chkVaporBTEX.Text = "BTEX"
        '
        'chkToCVapor
        '
        Me.chkToCVapor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkToCVapor.Location = New System.Drawing.Point(312, 24)
        Me.chkToCVapor.Name = "chkToCVapor"
        Me.chkToCVapor.Size = New System.Drawing.Size(72, 16)
        Me.chkToCVapor.TabIndex = 46
        Me.chkToCVapor.Text = "Vapor"
        '
        'chkFreeProductKerosene
        '
        Me.chkFreeProductKerosene.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFreeProductKerosene.Location = New System.Drawing.Point(216, 96)
        Me.chkFreeProductKerosene.Name = "chkFreeProductKerosene"
        Me.chkFreeProductKerosene.Size = New System.Drawing.Size(80, 16)
        Me.chkFreeProductKerosene.TabIndex = 43
        Me.chkFreeProductKerosene.Text = "Kerosene"
        '
        'chkFreeProductDiesel
        '
        Me.chkFreeProductDiesel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFreeProductDiesel.Location = New System.Drawing.Point(216, 72)
        Me.chkFreeProductDiesel.Name = "chkFreeProductDiesel"
        Me.chkFreeProductDiesel.Size = New System.Drawing.Size(64, 16)
        Me.chkFreeProductDiesel.TabIndex = 42
        Me.chkFreeProductDiesel.Text = "Diesel"
        '
        'chkFreeProductGasoline
        '
        Me.chkFreeProductGasoline.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFreeProductGasoline.Location = New System.Drawing.Point(216, 48)
        Me.chkFreeProductGasoline.Name = "chkFreeProductGasoline"
        Me.chkFreeProductGasoline.Size = New System.Drawing.Size(80, 16)
        Me.chkFreeProductGasoline.TabIndex = 41
        Me.chkFreeProductGasoline.Text = "Gasoline"
        '
        'chkToCFreeProduct
        '
        Me.chkToCFreeProduct.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkToCFreeProduct.Location = New System.Drawing.Point(200, 24)
        Me.chkToCFreeProduct.Name = "chkToCFreeProduct"
        Me.chkToCFreeProduct.Size = New System.Drawing.Size(96, 16)
        Me.chkToCFreeProduct.TabIndex = 40
        Me.chkToCFreeProduct.Text = "Free Product"
        '
        'chkGroundWaterTPH
        '
        Me.chkGroundWaterTPH.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGroundWaterTPH.Location = New System.Drawing.Point(104, 96)
        Me.chkGroundWaterTPH.Name = "chkGroundWaterTPH"
        Me.chkGroundWaterTPH.Size = New System.Drawing.Size(48, 16)
        Me.chkGroundWaterTPH.TabIndex = 39
        Me.chkGroundWaterTPH.Text = "TPH"
        '
        'chkGroundWaterPAH
        '
        Me.chkGroundWaterPAH.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGroundWaterPAH.Location = New System.Drawing.Point(104, 72)
        Me.chkGroundWaterPAH.Name = "chkGroundWaterPAH"
        Me.chkGroundWaterPAH.Size = New System.Drawing.Size(56, 16)
        Me.chkGroundWaterPAH.TabIndex = 38
        Me.chkGroundWaterPAH.Text = "PAH"
        '
        'chkGroundWaterBTEX
        '
        Me.chkGroundWaterBTEX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGroundWaterBTEX.Location = New System.Drawing.Point(104, 48)
        Me.chkGroundWaterBTEX.Name = "chkGroundWaterBTEX"
        Me.chkGroundWaterBTEX.Size = New System.Drawing.Size(56, 16)
        Me.chkGroundWaterBTEX.TabIndex = 37
        Me.chkGroundWaterBTEX.Text = "BTEX"
        '
        'chkGroundWater
        '
        Me.chkGroundWater.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGroundWater.Location = New System.Drawing.Point(88, 24)
        Me.chkGroundWater.Name = "chkGroundWater"
        Me.chkGroundWater.Size = New System.Drawing.Size(104, 16)
        Me.chkGroundWater.TabIndex = 36
        Me.chkGroundWater.Text = "Ground Water"
        '
        'chkSoilTPH
        '
        Me.chkSoilTPH.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSoilTPH.Location = New System.Drawing.Point(24, 96)
        Me.chkSoilTPH.Name = "chkSoilTPH"
        Me.chkSoilTPH.Size = New System.Drawing.Size(48, 16)
        Me.chkSoilTPH.TabIndex = 35
        Me.chkSoilTPH.Text = "TPH"
        '
        'chkSoilPAH
        '
        Me.chkSoilPAH.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSoilPAH.Location = New System.Drawing.Point(24, 72)
        Me.chkSoilPAH.Name = "chkSoilPAH"
        Me.chkSoilPAH.Size = New System.Drawing.Size(56, 16)
        Me.chkSoilPAH.TabIndex = 34
        Me.chkSoilPAH.Text = "PAH"
        '
        'chkSoilBTEX
        '
        Me.chkSoilBTEX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSoilBTEX.Location = New System.Drawing.Point(24, 48)
        Me.chkSoilBTEX.Name = "chkSoilBTEX"
        Me.chkSoilBTEX.Size = New System.Drawing.Size(56, 16)
        Me.chkSoilBTEX.TabIndex = 33
        Me.chkSoilBTEX.Text = "BTEX"
        '
        'chkSoil
        '
        Me.chkSoil.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSoil.Location = New System.Drawing.Point(8, 24)
        Me.chkSoil.Name = "chkSoil"
        Me.chkSoil.Size = New System.Drawing.Size(48, 16)
        Me.chkSoil.TabIndex = 32
        Me.chkSoil.Text = "Soil"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkInspection)
        Me.GroupBox1.Controls.Add(Me.chkTankClosure)
        Me.GroupBox1.Controls.Add(Me.chkInventoryShortage)
        Me.GroupBox1.Controls.Add(Me.chkFailedPTT)
        Me.GroupBox1.Controls.Add(Me.chkSoilContamination)
        Me.GroupBox1.Controls.Add(Me.chkFreeProduct)
        Me.GroupBox1.Controls.Add(Me.chkVapors)
        Me.GroupBox1.Controls.Add(Me.chkGWContamination)
        Me.GroupBox1.Controls.Add(Me.chkGWWell)
        Me.GroupBox1.Controls.Add(Me.chkSurfaceSheen)
        Me.GroupBox1.Controls.Add(Me.chkFacilityLeakDetection)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(24, 48)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(328, 168)
        Me.GroupBox1.TabIndex = 204
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "How Discovered"
        '
        'chkInspection
        '
        Me.chkInspection.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkInspection.Location = New System.Drawing.Point(176, 124)
        Me.chkInspection.Name = "chkInspection"
        Me.chkInspection.Size = New System.Drawing.Size(80, 16)
        Me.chkInspection.TabIndex = 19
        Me.chkInspection.Tag = "654"
        Me.chkInspection.Text = "Inspection"
        '
        'chkTankClosure
        '
        Me.chkTankClosure.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkTankClosure.Location = New System.Drawing.Point(176, 100)
        Me.chkTankClosure.Name = "chkTankClosure"
        Me.chkTankClosure.Size = New System.Drawing.Size(104, 16)
        Me.chkTankClosure.TabIndex = 18
        Me.chkTankClosure.Tag = "653"
        Me.chkTankClosure.Text = "Tank Closure"
        '
        'chkInventoryShortage
        '
        Me.chkInventoryShortage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkInventoryShortage.Location = New System.Drawing.Point(176, 76)
        Me.chkInventoryShortage.Name = "chkInventoryShortage"
        Me.chkInventoryShortage.Size = New System.Drawing.Size(128, 16)
        Me.chkInventoryShortage.TabIndex = 17
        Me.chkInventoryShortage.Tag = "652"
        Me.chkInventoryShortage.Text = "Inventory Shortage"
        '
        'chkFailedPTT
        '
        Me.chkFailedPTT.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFailedPTT.Location = New System.Drawing.Point(176, 52)
        Me.chkFailedPTT.Name = "chkFailedPTT"
        Me.chkFailedPTT.Size = New System.Drawing.Size(88, 16)
        Me.chkFailedPTT.TabIndex = 16
        Me.chkFailedPTT.Tag = "651"
        Me.chkFailedPTT.Text = "Failed PTT"
        '
        'chkSoilContamination
        '
        Me.chkSoilContamination.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSoilContamination.Location = New System.Drawing.Point(176, 28)
        Me.chkSoilContamination.Name = "chkSoilContamination"
        Me.chkSoilContamination.Size = New System.Drawing.Size(128, 16)
        Me.chkSoilContamination.TabIndex = 15
        Me.chkSoilContamination.Tag = "650"
        Me.chkSoilContamination.Text = "Soil Contamination"
        '
        'chkFreeProduct
        '
        Me.chkFreeProduct.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFreeProduct.Location = New System.Drawing.Point(8, 144)
        Me.chkFreeProduct.Name = "chkFreeProduct"
        Me.chkFreeProduct.Size = New System.Drawing.Size(96, 16)
        Me.chkFreeProduct.TabIndex = 11
        Me.chkFreeProduct.Tag = "649"
        Me.chkFreeProduct.Text = "Free Product"
        '
        'chkVapors
        '
        Me.chkVapors.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkVapors.Location = New System.Drawing.Point(8, 120)
        Me.chkVapors.Name = "chkVapors"
        Me.chkVapors.Size = New System.Drawing.Size(64, 16)
        Me.chkVapors.TabIndex = 10
        Me.chkVapors.Tag = "648"
        Me.chkVapors.Text = "Vapors"
        '
        'chkGWContamination
        '
        Me.chkGWContamination.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGWContamination.Location = New System.Drawing.Point(8, 96)
        Me.chkGWContamination.Name = "chkGWContamination"
        Me.chkGWContamination.Size = New System.Drawing.Size(128, 16)
        Me.chkGWContamination.TabIndex = 9
        Me.chkGWContamination.Tag = "647"
        Me.chkGWContamination.Text = "GW Contamination"
        '
        'chkGWWell
        '
        Me.chkGWWell.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGWWell.Location = New System.Drawing.Point(8, 72)
        Me.chkGWWell.Name = "chkGWWell"
        Me.chkGWWell.Size = New System.Drawing.Size(72, 16)
        Me.chkGWWell.TabIndex = 8
        Me.chkGWWell.Tag = "646"
        Me.chkGWWell.Text = "GW Well"
        '
        'chkSurfaceSheen
        '
        Me.chkSurfaceSheen.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSurfaceSheen.Location = New System.Drawing.Point(8, 48)
        Me.chkSurfaceSheen.Name = "chkSurfaceSheen"
        Me.chkSurfaceSheen.Size = New System.Drawing.Size(112, 16)
        Me.chkSurfaceSheen.TabIndex = 7
        Me.chkSurfaceSheen.Tag = "645"
        Me.chkSurfaceSheen.Text = "Surface Sheen"
        '
        'chkFacilityLeakDetection
        '
        Me.chkFacilityLeakDetection.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFacilityLeakDetection.Location = New System.Drawing.Point(8, 24)
        Me.chkFacilityLeakDetection.Name = "chkFacilityLeakDetection"
        Me.chkFacilityLeakDetection.Size = New System.Drawing.Size(152, 16)
        Me.chkFacilityLeakDetection.TabIndex = 6
        Me.chkFacilityLeakDetection.Tag = "644"
        Me.chkFacilityLeakDetection.Text = "Facility Leak Detection"
        '
        'ugTankandPipes
        '
        Me.ugTankandPipes.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTankandPipes.DisplayLayout.AutoFitColumns = True
        Me.ugTankandPipes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
        Me.ugTankandPipes.Location = New System.Drawing.Point(24, 224)
        Me.ugTankandPipes.Name = "ugTankandPipes"
        Me.ugTankandPipes.Size = New System.Drawing.Size(680, 192)
        Me.ugTankandPipes.TabIndex = 32
        Me.ugTankandPipes.Text = "Tank(s) / Pipe(s) Affected"
        '
        'cmbExtent
        '
        Me.cmbExtent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbExtent.DropDownWidth = 300
        Me.cmbExtent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbExtent.ItemHeight = 15
        Me.cmbExtent.Location = New System.Drawing.Point(638, 8)
        Me.cmbExtent.Name = "cmbExtent"
        Me.cmbExtent.Size = New System.Drawing.Size(128, 23)
        Me.cmbExtent.TabIndex = 3
        '
        'lblExtent
        '
        Me.lblExtent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExtent.Location = New System.Drawing.Point(593, 8)
        Me.lblExtent.Name = "lblExtent"
        Me.lblExtent.Size = New System.Drawing.Size(48, 23)
        Me.lblExtent.TabIndex = 187
        Me.lblExtent.Text = "Extent:"
        Me.lblExtent.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmbLocation
        '
        Me.cmbLocation.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLocation.DropDownWidth = 125
        Me.cmbLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbLocation.ItemHeight = 15
        Me.cmbLocation.Location = New System.Drawing.Point(461, 8)
        Me.cmbLocation.Name = "cmbLocation"
        Me.cmbLocation.Size = New System.Drawing.Size(128, 23)
        Me.cmbLocation.TabIndex = 2
        '
        'cmbIdentifiedBy
        '
        Me.cmbIdentifiedBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbIdentifiedBy.DropDownWidth = 125
        Me.cmbIdentifiedBy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbIdentifiedBy.ItemHeight = 15
        Me.cmbIdentifiedBy.Location = New System.Drawing.Point(280, 8)
        Me.cmbIdentifiedBy.Name = "cmbIdentifiedBy"
        Me.cmbIdentifiedBy.Size = New System.Drawing.Size(128, 23)
        Me.cmbIdentifiedBy.TabIndex = 1
        '
        'lblLocation
        '
        Me.lblLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLocation.Location = New System.Drawing.Point(413, 8)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(48, 23)
        Me.lblLocation.TabIndex = 178
        Me.lblLocation.Text = "Source:"
        Me.lblLocation.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblIdentifiedBy
        '
        Me.lblIdentifiedBy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIdentifiedBy.Location = New System.Drawing.Point(205, 8)
        Me.lblIdentifiedBy.Name = "lblIdentifiedBy"
        Me.lblIdentifiedBy.Size = New System.Drawing.Size(76, 23)
        Me.lblIdentifiedBy.TabIndex = 176
        Me.lblIdentifiedBy.Text = "Identified By:"
        Me.lblIdentifiedBy.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblConfirmedOn
        '
        Me.lblConfirmedOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConfirmedOn.Location = New System.Drawing.Point(8, 8)
        Me.lblConfirmedOn.Name = "lblConfirmedOn"
        Me.lblConfirmedOn.Size = New System.Drawing.Size(88, 17)
        Me.lblConfirmedOn.TabIndex = 137
        Me.lblConfirmedOn.Text = "Confirmed On:"
        Me.lblConfirmedOn.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbCause
        '
        Me.cmbCause.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCause.DropDownWidth = 300
        Me.cmbCause.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCause.ItemHeight = 15
        Me.cmbCause.Location = New System.Drawing.Point(815, 8)
        Me.cmbCause.Name = "cmbCause"
        Me.cmbCause.Size = New System.Drawing.Size(128, 23)
        Me.cmbCause.TabIndex = 3
        '
        'lblCause
        '
        Me.lblCause.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCause.Location = New System.Drawing.Point(769, 8)
        Me.lblCause.Name = "lblCause"
        Me.lblCause.Size = New System.Drawing.Size(48, 23)
        Me.lblCause.TabIndex = 187
        Me.lblCause.Text = "Cause:"
        Me.lblCause.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'pnlReleaseInfo
        '
        Me.pnlReleaseInfo.Controls.Add(Me.lblReleaseInfoHead)
        Me.pnlReleaseInfo.Controls.Add(Me.lblReleaseInfoDisplay)
        Me.pnlReleaseInfo.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlReleaseInfo.Location = New System.Drawing.Point(0, 152)
        Me.pnlReleaseInfo.Name = "pnlReleaseInfo"
        Me.pnlReleaseInfo.Size = New System.Drawing.Size(996, 24)
        Me.pnlReleaseInfo.TabIndex = 11
        '
        'lblReleaseInfoHead
        '
        Me.lblReleaseInfoHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblReleaseInfoHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblReleaseInfoHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblReleaseInfoHead.Location = New System.Drawing.Point(16, 0)
        Me.lblReleaseInfoHead.Name = "lblReleaseInfoHead"
        Me.lblReleaseInfoHead.Size = New System.Drawing.Size(980, 24)
        Me.lblReleaseInfoHead.TabIndex = 1
        Me.lblReleaseInfoHead.Text = "Release Info"
        Me.lblReleaseInfoHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblReleaseInfoDisplay
        '
        Me.lblReleaseInfoDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblReleaseInfoDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblReleaseInfoDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblReleaseInfoDisplay.Name = "lblReleaseInfoDisplay"
        Me.lblReleaseInfoDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblReleaseInfoDisplay.TabIndex = 0
        Me.lblReleaseInfoDisplay.Text = "-"
        Me.lblReleaseInfoDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PnlEventInfoDetails
        '
        Me.PnlEventInfoDetails.AutoScroll = True
        Me.PnlEventInfoDetails.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PnlEventInfoDetails.Controls.Add(Me.lblCompAssDate)
        Me.PnlEventInfoDetails.Controls.Add(Me.dtCompAssDate)
        Me.PnlEventInfoDetails.Controls.Add(Me.dtLastGWS)
        Me.PnlEventInfoDetails.Controls.Add(Me.dtLastPTT)
        Me.PnlEventInfoDetails.Controls.Add(Me.dtLastLDR)
        Me.PnlEventInfoDetails.Controls.Add(Me.dtStartDate)
        Me.PnlEventInfoDetails.Controls.Add(Me.lblLastGWS)
        Me.PnlEventInfoDetails.Controls.Add(Me.txtRelatedSites)
        Me.PnlEventInfoDetails.Controls.Add(Me.lblRelatedSites)
        Me.PnlEventInfoDetails.Controls.Add(Me.cmbReleaseStatus)
        Me.PnlEventInfoDetails.Controls.Add(Me.cmbSuspectedSource)
        Me.PnlEventInfoDetails.Controls.Add(Me.lblSuspectedSource)
        Me.PnlEventInfoDetails.Controls.Add(Me.lblLastPTT)
        Me.PnlEventInfoDetails.Controls.Add(Me.lblLastLDR)
        Me.PnlEventInfoDetails.Controls.Add(Me.lblReleaseStatus)
        Me.PnlEventInfoDetails.Controls.Add(Me.lblDateofReport)
        Me.PnlEventInfoDetails.Controls.Add(Me.dtDateofReport)
        Me.PnlEventInfoDetails.Controls.Add(Me.lblEventStartDate)
        Me.PnlEventInfoDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.PnlEventInfoDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlEventInfoDetails.Location = New System.Drawing.Point(0, 24)
        Me.PnlEventInfoDetails.Name = "PnlEventInfoDetails"
        Me.PnlEventInfoDetails.Size = New System.Drawing.Size(996, 128)
        Me.PnlEventInfoDetails.TabIndex = 10
        '
        'lblCompAssDate
        '
        Me.lblCompAssDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompAssDate.Location = New System.Drawing.Point(16, 76)
        Me.lblCompAssDate.Name = "lblCompAssDate"
        Me.lblCompAssDate.Size = New System.Drawing.Size(96, 32)
        Me.lblCompAssDate.TabIndex = 211
        Me.lblCompAssDate.Text = "Site Completely Assessed Date:"
        '
        'dtCompAssDate
        '
        Me.dtCompAssDate.Checked = False
        Me.dtCompAssDate.CustomFormat = "__/__/____"
        Me.dtCompAssDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtCompAssDate.Location = New System.Drawing.Point(112, 88)
        Me.dtCompAssDate.Name = "dtCompAssDate"
        Me.dtCompAssDate.ShowCheckBox = True
        Me.dtCompAssDate.Size = New System.Drawing.Size(104, 21)
        Me.dtCompAssDate.TabIndex = 210
        '
        'dtLastGWS
        '
        Me.dtLastGWS.Checked = False
        Me.dtLastGWS.CustomFormat = "__/__/____"
        Me.dtLastGWS.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtLastGWS.Location = New System.Drawing.Point(664, 80)
        Me.dtLastGWS.Name = "dtLastGWS"
        Me.dtLastGWS.ShowCheckBox = True
        Me.dtLastGWS.Size = New System.Drawing.Size(104, 21)
        Me.dtLastGWS.TabIndex = 209
        '
        'dtLastPTT
        '
        Me.dtLastPTT.Checked = False
        Me.dtLastPTT.CustomFormat = "__/__/____"
        Me.dtLastPTT.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtLastPTT.Location = New System.Drawing.Point(664, 48)
        Me.dtLastPTT.Name = "dtLastPTT"
        Me.dtLastPTT.ShowCheckBox = True
        Me.dtLastPTT.Size = New System.Drawing.Size(104, 21)
        Me.dtLastPTT.TabIndex = 208
        '
        'dtLastLDR
        '
        Me.dtLastLDR.Checked = False
        Me.dtLastLDR.CustomFormat = "__/__/____"
        Me.dtLastLDR.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtLastLDR.Location = New System.Drawing.Point(664, 16)
        Me.dtLastLDR.Name = "dtLastLDR"
        Me.dtLastLDR.ShowCheckBox = True
        Me.dtLastLDR.Size = New System.Drawing.Size(104, 21)
        Me.dtLastLDR.TabIndex = 207
        '
        'dtStartDate
        '
        Me.dtStartDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtStartDate.Checked = False
        Me.dtStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtStartDate.Location = New System.Drawing.Point(112, 16)
        Me.dtStartDate.Name = "dtStartDate"
        Me.dtStartDate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.dtStartDate.Size = New System.Drawing.Size(104, 21)
        Me.dtStartDate.TabIndex = 188
        Me.dtStartDate.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'lblLastGWS
        '
        Me.lblLastGWS.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastGWS.Location = New System.Drawing.Point(592, 80)
        Me.lblLastGWS.Name = "lblLastGWS"
        Me.lblLastGWS.Size = New System.Drawing.Size(72, 23)
        Me.lblLastGWS.TabIndex = 187
        Me.lblLastGWS.Text = "Last GWS:"
        Me.lblLastGWS.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'txtRelatedSites
        '
        Me.txtRelatedSites.AcceptsTab = True
        Me.txtRelatedSites.AutoSize = False
        Me.txtRelatedSites.Location = New System.Drawing.Point(336, 80)
        Me.txtRelatedSites.Name = "txtRelatedSites"
        Me.txtRelatedSites.Size = New System.Drawing.Size(256, 23)
        Me.txtRelatedSites.TabIndex = 4
        Me.txtRelatedSites.Text = ""
        Me.txtRelatedSites.WordWrap = False
        '
        'lblRelatedSites
        '
        Me.lblRelatedSites.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRelatedSites.Location = New System.Drawing.Point(248, 80)
        Me.lblRelatedSites.Name = "lblRelatedSites"
        Me.lblRelatedSites.Size = New System.Drawing.Size(88, 23)
        Me.lblRelatedSites.TabIndex = 185
        Me.lblRelatedSites.Text = "Related Sites:"
        Me.lblRelatedSites.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbReleaseStatus
        '
        Me.cmbReleaseStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbReleaseStatus.DropDownWidth = 140
        Me.cmbReleaseStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbReleaseStatus.ItemHeight = 15
        Me.cmbReleaseStatus.Location = New System.Drawing.Point(336, 16)
        Me.cmbReleaseStatus.Name = "cmbReleaseStatus"
        Me.cmbReleaseStatus.Size = New System.Drawing.Size(128, 23)
        Me.cmbReleaseStatus.TabIndex = 2
        '
        'cmbSuspectedSource
        '
        Me.cmbSuspectedSource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSuspectedSource.DropDownWidth = 300
        Me.cmbSuspectedSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSuspectedSource.ItemHeight = 15
        Me.cmbSuspectedSource.Location = New System.Drawing.Point(336, 48)
        Me.cmbSuspectedSource.Name = "cmbSuspectedSource"
        Me.cmbSuspectedSource.Size = New System.Drawing.Size(240, 23)
        Me.cmbSuspectedSource.TabIndex = 3
        '
        'lblSuspectedSource
        '
        Me.lblSuspectedSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSuspectedSource.Location = New System.Drawing.Point(224, 48)
        Me.lblSuspectedSource.Name = "lblSuspectedSource"
        Me.lblSuspectedSource.Size = New System.Drawing.Size(112, 17)
        Me.lblSuspectedSource.TabIndex = 183
        Me.lblSuspectedSource.Text = "Suspected Source:"
        Me.lblSuspectedSource.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLastPTT
        '
        Me.lblLastPTT.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastPTT.Location = New System.Drawing.Point(592, 48)
        Me.lblLastPTT.Name = "lblLastPTT"
        Me.lblLastPTT.Size = New System.Drawing.Size(72, 23)
        Me.lblLastPTT.TabIndex = 180
        Me.lblLastPTT.Text = "Last PTT:"
        Me.lblLastPTT.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblLastLDR
        '
        Me.lblLastLDR.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLastLDR.Location = New System.Drawing.Point(592, 16)
        Me.lblLastLDR.Name = "lblLastLDR"
        Me.lblLastLDR.Size = New System.Drawing.Size(72, 23)
        Me.lblLastLDR.TabIndex = 178
        Me.lblLastLDR.Text = "Last LDR:"
        Me.lblLastLDR.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblReleaseStatus
        '
        Me.lblReleaseStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReleaseStatus.Location = New System.Drawing.Point(240, 16)
        Me.lblReleaseStatus.Name = "lblReleaseStatus"
        Me.lblReleaseStatus.Size = New System.Drawing.Size(96, 23)
        Me.lblReleaseStatus.TabIndex = 176
        Me.lblReleaseStatus.Text = "Release Status:"
        Me.lblReleaseStatus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblDateofReport
        '
        Me.lblDateofReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateofReport.Location = New System.Drawing.Point(16, 48)
        Me.lblDateofReport.Name = "lblDateofReport"
        Me.lblDateofReport.Size = New System.Drawing.Size(88, 17)
        Me.lblDateofReport.TabIndex = 174
        Me.lblDateofReport.Text = "Date of Report:"
        '
        'dtDateofReport
        '
        Me.dtDateofReport.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtDateofReport.Checked = False
        Me.dtDateofReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtDateofReport.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtDateofReport.Location = New System.Drawing.Point(112, 48)
        Me.dtDateofReport.Name = "dtDateofReport"
        Me.dtDateofReport.ShowCheckBox = True
        Me.dtDateofReport.Size = New System.Drawing.Size(104, 21)
        Me.dtDateofReport.TabIndex = 1
        Me.dtDateofReport.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'lblEventStartDate
        '
        Me.lblEventStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEventStartDate.Location = New System.Drawing.Point(32, 16)
        Me.lblEventStartDate.Name = "lblEventStartDate"
        Me.lblEventStartDate.Size = New System.Drawing.Size(71, 17)
        Me.lblEventStartDate.TabIndex = 137
        Me.lblEventStartDate.Text = "Start Date:"
        Me.lblEventStartDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlEventInfo
        '
        Me.pnlEventInfo.Controls.Add(Me.lblEventInfoHead)
        Me.pnlEventInfo.Controls.Add(Me.lblEventInfoDisplay)
        Me.pnlEventInfo.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlEventInfo.Location = New System.Drawing.Point(0, 0)
        Me.pnlEventInfo.Name = "pnlEventInfo"
        Me.pnlEventInfo.Size = New System.Drawing.Size(996, 24)
        Me.pnlEventInfo.TabIndex = 9
        '
        'lblEventInfoHead
        '
        Me.lblEventInfoHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblEventInfoHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblEventInfoHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblEventInfoHead.Location = New System.Drawing.Point(16, 0)
        Me.lblEventInfoHead.Name = "lblEventInfoHead"
        Me.lblEventInfoHead.Size = New System.Drawing.Size(980, 24)
        Me.lblEventInfoHead.TabIndex = 1
        Me.lblEventInfoHead.Text = "Event Info"
        Me.lblEventInfoHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEventInfoDisplay
        '
        Me.lblEventInfoDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEventInfoDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblEventInfoDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblEventInfoDisplay.Name = "lblEventInfoDisplay"
        Me.lblEventInfoDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblEventInfoDisplay.TabIndex = 0
        Me.lblEventInfoDisplay.Text = "-"
        Me.lblEventInfoDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlLUSTEventHeader
        '
        Me.pnlLUSTEventHeader.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblPMHistorytt)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblEventIDValue)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblPrioritytt)
        Me.pnlLUSTEventHeader.Controls.Add(Me.cmbPriority)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblPriority)
        Me.pnlLUSTEventHeader.Controls.Add(Me.cmbMGPTFStatus)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblMGPTFStatus)
        Me.pnlLUSTEventHeader.Controls.Add(Me.cmbEventStatus)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblEventStatus)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblEventCountValue)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblProject234)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblEventID)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lblProjectManager)
        Me.pnlLUSTEventHeader.Controls.Add(Me.cmbProjectManager)
        Me.pnlLUSTEventHeader.Controls.Add(Me.lnkEnsite)
        Me.pnlLUSTEventHeader.Controls.Add(Me.btnGoToFinancial)
        Me.pnlLUSTEventHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLUSTEventHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlLUSTEventHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlLUSTEventHeader.Name = "pnlLUSTEventHeader"
        Me.pnlLUSTEventHeader.Size = New System.Drawing.Size(1016, 56)
        Me.pnlLUSTEventHeader.TabIndex = 5
        '
        'lblPMHistorytt
        '
        Me.lblPMHistorytt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPMHistorytt.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPMHistorytt.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblPMHistorytt.Location = New System.Drawing.Point(312, 16)
        Me.lblPMHistorytt.Name = "lblPMHistorytt"
        Me.lblPMHistorytt.Size = New System.Drawing.Size(16, 23)
        Me.lblPMHistorytt.TabIndex = 218
        Me.lblPMHistorytt.Text = "H"
        Me.lblPMHistorytt.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblEventIDValue
        '
        Me.lblEventIDValue.AutoSize = True
        Me.lblEventIDValue.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblEventIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEventIDValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblEventIDValue.Location = New System.Drawing.Point(64, 17)
        Me.lblEventIDValue.Name = "lblEventIDValue"
        Me.lblEventIDValue.Size = New System.Drawing.Size(0, 16)
        Me.lblEventIDValue.TabIndex = 209
        Me.lblEventIDValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPrioritytt
        '
        Me.lblPrioritytt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrioritytt.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrioritytt.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblPrioritytt.Location = New System.Drawing.Point(848, 16)
        Me.lblPrioritytt.Name = "lblPrioritytt"
        Me.lblPrioritytt.Size = New System.Drawing.Size(16, 23)
        Me.lblPrioritytt.TabIndex = 217
        Me.lblPrioritytt.Text = "i"
        Me.lblPrioritytt.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'cmbPriority
        '
        Me.cmbPriority.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPriority.DropDownWidth = 88
        Me.cmbPriority.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPriority.ItemHeight = 15
        Me.cmbPriority.Location = New System.Drawing.Point(760, 16)
        Me.cmbPriority.MaxDropDownItems = 80
        Me.cmbPriority.Name = "cmbPriority"
        Me.cmbPriority.Size = New System.Drawing.Size(88, 23)
        Me.cmbPriority.TabIndex = 3
        '
        'lblPriority
        '
        Me.lblPriority.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPriority.Location = New System.Drawing.Point(696, 16)
        Me.lblPriority.Name = "lblPriority"
        Me.lblPriority.Size = New System.Drawing.Size(56, 17)
        Me.lblPriority.TabIndex = 216
        Me.lblPriority.Text = "Priority"
        Me.lblPriority.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbMGPTFStatus
        '
        Me.cmbMGPTFStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMGPTFStatus.DropDownWidth = 100
        Me.cmbMGPTFStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbMGPTFStatus.ItemHeight = 15
        Me.cmbMGPTFStatus.Location = New System.Drawing.Point(600, 16)
        Me.cmbMGPTFStatus.MaxDropDownItems = 80
        Me.cmbMGPTFStatus.Name = "cmbMGPTFStatus"
        Me.cmbMGPTFStatus.Size = New System.Drawing.Size(96, 23)
        Me.cmbMGPTFStatus.TabIndex = 2
        '
        'lblMGPTFStatus
        '
        Me.lblMGPTFStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMGPTFStatus.Location = New System.Drawing.Point(504, 16)
        Me.lblMGPTFStatus.Name = "lblMGPTFStatus"
        Me.lblMGPTFStatus.Size = New System.Drawing.Size(96, 17)
        Me.lblMGPTFStatus.TabIndex = 214
        Me.lblMGPTFStatus.Text = "MGPTF Status"
        Me.lblMGPTFStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbEventStatus
        '
        Me.cmbEventStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbEventStatus.DropDownWidth = 88
        Me.cmbEventStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbEventStatus.ItemHeight = 15
        Me.cmbEventStatus.Location = New System.Drawing.Point(408, 16)
        Me.cmbEventStatus.MaxDropDownItems = 80
        Me.cmbEventStatus.Name = "cmbEventStatus"
        Me.cmbEventStatus.Size = New System.Drawing.Size(88, 23)
        Me.cmbEventStatus.TabIndex = 1
        '
        'lblEventStatus
        '
        Me.lblEventStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEventStatus.Location = New System.Drawing.Point(328, 16)
        Me.lblEventStatus.Name = "lblEventStatus"
        Me.lblEventStatus.Size = New System.Drawing.Size(80, 17)
        Me.lblEventStatus.TabIndex = 212
        Me.lblEventStatus.Text = "Event Status:"
        Me.lblEventStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblEventCountValue
        '
        Me.lblEventCountValue.AutoSize = True
        Me.lblEventCountValue.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblEventCountValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEventCountValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblEventCountValue.Location = New System.Drawing.Point(80, 16)
        Me.lblEventCountValue.Name = "lblEventCountValue"
        Me.lblEventCountValue.Size = New System.Drawing.Size(0, 17)
        Me.lblEventCountValue.TabIndex = 210
        '
        'lblProject234
        '
        Me.lblProject234.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProject234.Location = New System.Drawing.Point(120, 16)
        Me.lblProject234.Name = "lblProject234"
        Me.lblProject234.Size = New System.Drawing.Size(72, 17)
        Me.lblProject234.TabIndex = 139
        Me.lblProject234.Text = "Project Mgr"
        Me.lblProject234.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblEventID
        '
        Me.lblEventID.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblEventID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEventID.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblEventID.Location = New System.Drawing.Point(8, 16)
        Me.lblEventID.Name = "lblEventID"
        Me.lblEventID.Size = New System.Drawing.Size(120, 17)
        Me.lblEventID.TabIndex = 208
        Me.lblEventID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblProjectManager
        '
        Me.lblProjectManager.BackColor = System.Drawing.SystemColors.Window
        Me.lblProjectManager.Location = New System.Drawing.Point(200, 16)
        Me.lblProjectManager.Name = "lblProjectManager"
        Me.lblProjectManager.Size = New System.Drawing.Size(112, 17)
        Me.lblProjectManager.TabIndex = 220
        Me.lblProjectManager.Text = "Label2"
        Me.lblProjectManager.Visible = False
        '
        'cmbProjectManager
        '
        Me.cmbProjectManager.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbProjectManager.DropDownWidth = 300
        Me.cmbProjectManager.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbProjectManager.ItemHeight = 15
        Me.cmbProjectManager.Location = New System.Drawing.Point(200, 16)
        Me.cmbProjectManager.Name = "cmbProjectManager"
        Me.cmbProjectManager.Size = New System.Drawing.Size(112, 23)
        Me.cmbProjectManager.TabIndex = 0
        '
        'lnkEnsite
        '
        Me.lnkEnsite.Location = New System.Drawing.Point(872, 32)
        Me.lnkEnsite.Name = "lnkEnsite"
        Me.lnkEnsite.Size = New System.Drawing.Size(136, 16)
        Me.lnkEnsite.TabIndex = 1
        Me.lnkEnsite.TabStop = True
        Me.lnkEnsite.Text = "Access Ensite Permits"
        '
        'btnGoToFinancial
        '
        Me.btnGoToFinancial.Location = New System.Drawing.Point(872, 8)
        Me.btnGoToFinancial.Name = "btnGoToFinancial"
        Me.btnGoToFinancial.Size = New System.Drawing.Size(96, 23)
        Me.btnGoToFinancial.TabIndex = 221
        Me.btnGoToFinancial.Text = "View Financial"
        '
        'pnlLUSTEventBottom
        '
        Me.pnlLUSTEventBottom.Controls.Add(Me.btnActPlanning)
        Me.pnlLUSTEventBottom.Controls.Add(Me.btnCancel)
        Me.pnlLUSTEventBottom.Controls.Add(Me.btnDeleteLUSTEvent)
        Me.pnlLUSTEventBottom.Controls.Add(Me.btnSaveLUSTEvent)
        Me.pnlLUSTEventBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLUSTEventBottom.Location = New System.Drawing.Point(0, 604)
        Me.pnlLUSTEventBottom.Name = "pnlLUSTEventBottom"
        Me.pnlLUSTEventBottom.Size = New System.Drawing.Size(1016, 40)
        Me.pnlLUSTEventBottom.TabIndex = 3
        '
        'btnActPlanning
        '
        Me.btnActPlanning.Enabled = False
        Me.btnActPlanning.Location = New System.Drawing.Point(864, 8)
        Me.btnActPlanning.Name = "btnActPlanning"
        Me.btnActPlanning.Size = New System.Drawing.Size(128, 23)
        Me.btnActPlanning.TabIndex = 7
        Me.btnActPlanning.Text = "Plan Activities"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(600, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(120, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        '
        'btnDeleteLUSTEvent
        '
        Me.btnDeleteLUSTEvent.Enabled = False
        Me.btnDeleteLUSTEvent.Location = New System.Drawing.Point(728, 8)
        Me.btnDeleteLUSTEvent.Name = "btnDeleteLUSTEvent"
        Me.btnDeleteLUSTEvent.Size = New System.Drawing.Size(128, 23)
        Me.btnDeleteLUSTEvent.TabIndex = 5
        Me.btnDeleteLUSTEvent.Text = "Delete LUST Event"
        '
        'btnSaveLUSTEvent
        '
        Me.btnSaveLUSTEvent.Location = New System.Drawing.Point(480, 8)
        Me.btnSaveLUSTEvent.Name = "btnSaveLUSTEvent"
        Me.btnSaveLUSTEvent.Size = New System.Drawing.Size(112, 23)
        Me.btnSaveLUSTEvent.TabIndex = 4
        Me.btnSaveLUSTEvent.Text = "Save LUST Event"
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
        Me.tbPageSummary.Size = New System.Drawing.Size(1016, 644)
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
        Me.pnlOwnerSummaryDetails.Size = New System.Drawing.Size(1004, 624)
        Me.pnlOwnerSummaryDetails.TabIndex = 7
        '
        'UCOwnerSummary
        '
        Me.UCOwnerSummary.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerSummary.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerSummary.Name = "UCOwnerSummary"
        Me.UCOwnerSummary.Size = New System.Drawing.Size(1004, 624)
        Me.UCOwnerSummary.TabIndex = 0
        '
        'Panel12
        '
        Me.Panel12.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel12.DockPadding.Left = 10
        Me.Panel12.Location = New System.Drawing.Point(1004, 16)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(8, 624)
        Me.Panel12.TabIndex = 6
        '
        'pnlOwnerSummaryHeader
        '
        Me.pnlOwnerSummaryHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerSummaryHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlOwnerSummaryHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerSummaryHeader.Name = "pnlOwnerSummaryHeader"
        Me.pnlOwnerSummaryHeader.Size = New System.Drawing.Size(1012, 16)
        Me.pnlOwnerSummaryHeader.TabIndex = 2
        '
        'Technical
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1024, 694)
        Me.Controls.Add(Me.tbCntrlTechnical)
        Me.Controls.Add(Me.pnlTop)
        Me.Name = "Technical"
        Me.Text = "Technical"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTop.ResumeLayout(False)
        Me.tbCntrlTechnical.ResumeLayout(False)
        Me.tbPageOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerBottom.ResumeLayout(False)
        Me.tbCtrlOwner.ResumeLayout(False)
        Me.tbPageOwnerFacilities.ResumeLayout(False)
        CType(Me.ugFacilityList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerFacilityBottom.ResumeLayout(False)
        Me.tbPageOwnerContactList.ResumeLayout(False)
        Me.pnlOwnerContactContainer.ResumeLayout(False)
        CType(Me.ugOwnerContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerContactHeader.ResumeLayout(False)
        Me.pnlOwnerContactButtons.ResumeLayout(False)
        Me.tbPageOwnerDocuments.ResumeLayout(False)
        Me.pnlOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerButtons.ResumeLayout(False)
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageFacilityDetail.ResumeLayout(False)
        Me.pnlFacilityBottom.ResumeLayout(False)
        Me.tbCntrlFacility.ResumeLayout(False)
        Me.tbPageAddLustEvent.ResumeLayout(False)
        CType(Me.dgLUSTEvents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityLustButton.ResumeLayout(False)
        Me.tbPageFacilityDocuments.ResumeLayout(False)
        Me.pnl_FacilityDetail.ResumeLayout(False)
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageLUSTEvent.ResumeLayout(False)
        Me.pnlLustEvents.ResumeLayout(False)
        Me.pnlLUSTEventDetails.ResumeLayout(False)
        Me.pnlERACandContactsDetails.ResumeLayout(False)
        Me.pnlERACContactButtons.ResumeLayout(False)
        Me.pnlERACContactContainer.ResumeLayout(False)
        CType(Me.ugERACContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlERACContactHeader.ResumeLayout(False)
        Me.pnlERACandIRAC.ResumeLayout(False)
        Me.pnlERACandContacts.ResumeLayout(False)
        Me.pnlRemediationSystemsDetails.ResumeLayout(False)
        CType(Me.ugRemediationSystem, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlRemediationSystems.ResumeLayout(False)
        Me.pnlCommentsDetails.ResumeLayout(False)
        Me.pnlComments.ResumeLayout(False)
        Me.pnlActivitiesDocumentsDetails.ResumeLayout(False)
        CType(Me.ugActivitiesandDocuments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlActivitiesDocuments.ResumeLayout(False)
        Me.pnlFundsEligibilityDetails.ResumeLayout(False)
        Me.pnlFundsEligibilityQuestions.ResumeLayout(False)
        Me.pnlTFAssess.ResumeLayout(False)
        Me.pnlFundsEligibilitySysQuestions.ResumeLayout(False)
        Me.pnlFundsEligibility.ResumeLayout(False)
        Me.PnlReleaseInfoDetails.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.ugTankandPipes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlReleaseInfo.ResumeLayout(False)
        Me.PnlEventInfoDetails.ResumeLayout(False)
        Me.pnlEventInfo.ResumeLayout(False)
        Me.pnlLUSTEventHeader.ResumeLayout(False)
        Me.pnlLUSTEventBottom.ResumeLayout(False)
        Me.tbPageSummary.ResumeLayout(False)
        Me.pnlOwnerSummaryDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Intialization"
    Private Sub InitControls()
        UIUtilsGen.PopulateOwnerType(cmbOwnerType, oOwner)
        'If oOwner.ID <> 0 Then
        If oOwner.ID >= 0 Then
            UIUtilsGen.PopulateFacilityType(Me.cmbFacilityType, oOwner.Facilities)
            UIUtilsGen.PopulateFacilityDatum(Me.cmbFacilityDatum, oOwner.Facilities)
            UIUtilsGen.PopulateFacilityMethod(Me.cmbFacilityMethod, oOwner.Facilities)
            UIUtilsGen.PopulateFacilityLocationType(Me.cmbFacilityLocationType, oOwner.Facilities)
        End If
        btnFacilitySave.Enabled = False
        btnFacilityCancel.Enabled = False
    End Sub
    Private Sub Technical_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim oProfile As New MUSTER.BusinessLogic.pProfile
        Dim strProfileKey As String

        dtPickUpcomingInstallDateValue.Enabled = False

        bolLoading = False
        SetToolTips()

        If PMHead_StaffID = 0 Then
            Dim oUser As New MUSTER.BusinessLogic.pUser
            Dim oUserInfo As New MUSTER.Info.UserInfo
            oUserInfo = oUser.RetrievePMHead
            PMHead_UserID = oUserInfo.ID
            PMHead_StaffID = oUserInfo.UserKey
            UserID = MusterContainer.AppUser.ID
            UserStaffID = MusterContainer.AppUser.UserKey
            If UserID.ToUpper = PMHead_UserID.ToUpper Then
                bolPMHead = True
            Else
                bolPMHead = False
            End If
        End If

        SetPageSecurity(False)

        If tbCntrlTechnical.SelectedTab.Name = tbPageLUSTEvent.Name Then
            btnViewModifyComment.Select()
            btnViewModifyComment.Focus()
        End If
    End Sub
    Private Sub SetPageSecurity(ByVal AddNew As Boolean)
        Dim PMDescription As String = String.Empty

        If oLustEvent.PMDesc Is Nothing Then
            PMDescription = String.Empty
        Else
            PMDescription = oLustEvent.PMDesc.ToUpper
        End If

        If Not AddNew Then
            Me.btnActPlanning.Enabled = True
        End If

        If MusterContainer.AppUser.DefaultModule = 616 Then

            Text = "Financial View of Tenchical Data"
            pnlActivitiesDocumentsDetails.Visible = False
            pnlCommentsDetails.Visible = False
            Me.pnlERACandContactsDetails.Visible = True

            BtnSaveEngineers.Enabled = True
            pnlLUSTEventHeader.Enabled = False
            lblEventInfoDisplay.Enabled = False
            lblActivitiesDocumentsDisplay.Enabled = False
            lblReleaseInfoDisplay.Enabled = False
            lblRemediationSystemsDisplay.Enabled = False
            lblCommentsDisplay.Enabled = False
            lblERACandContactsDisplay.Enabled = False
            lblFundsEligibilityDisplay.Enabled = False

        End If

        If AddNew Or (bolPMHead Or MusterContainer.AppUser.HEAD_ADMIN) Or PMDescription.ToUpper = MusterContainer.AppUser.ID.ToUpper Or cmbProjectManager.Text.ToUpper = MusterContainer.AppUser.Name.ToUpper Then
            btnDeleteActivity.Visible = True
            btnDeleteDocument.Visible = True

            'pnlTFAssess.Enabled = True
            'pnlActivitiesDocumentsDetails.Enabled = True
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow = ugActivitiesandDocuments.ActiveRow

            If Not ugRow Is Nothing Then
                If ugRow.Band.Index = 1 Then
                    ugRow = ugRow.ParentRow
                End If
                If ugRow.Cells("ActivityDocsRelCount").Value > 0 Then
                    btnAddDocument.Enabled = True
                    btnModifyDocument.Enabled = True
                    btnDeleteDocument.Enabled = True
                End If
            End If

            btnAddActivity.Enabled = True
            btnModifyActivity.Enabled = True
            'btnAddDocument.Enabled = True
            'btnModifyDocument.Enabled = True
            btnDeleteActivity.Enabled = True
            'btnDeleteDocument.Enabled = True
            btnSaveLUSTEvent.Enabled = True
            bolAllowDocPop = True
        Else
            btnDeleteActivity.Visible = False
            btnDeleteDocument.Visible = False
            'pnlTFAssess.Enabled = False
            'pnlActivitiesDocumentsDetails.Enabled = False
            btnAddActivity.Enabled = False
            btnModifyActivity.Enabled = False
            btnAddDocument.Enabled = False
            btnModifyDocument.Enabled = False
            btnDeleteActivity.Enabled = False
            btnDeleteDocument.Enabled = False
            btnSaveLUSTEvent.Enabled = False
            bolAllowDocPop = False
        End If
        If (bolPMHead Or MusterContainer.AppUser.HEAD_ADMIN) Then
            pnlTFAssess.Enabled = True
            btnDeleteLUSTEvent.Visible = True
            btnDeleteLUSTEvent.Enabled = True
            btnTransferLUSTEvent.Visible = True
        Else
            pnlTFAssess.Enabled = False
            btnDeleteLUSTEvent.Visible = False
            btnDeleteLUSTEvent.Enabled = False
            btnTransferLUSTEvent.Visible = False
        End If
    End Sub
    Private Sub Technical_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Technical")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Property FormLoading() As Boolean
        Get
            Return bolLoading
        End Get
        Set(ByVal Value As Boolean)
            bolLoading = Value
        End Set
    End Property
    Private Sub Technical_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Technical")
            If lblOwnerIDValue.Text <> String.Empty Then
                If oOwner.ID <> lblOwnerIDValue.Text Then
                    oOwner.Retrieve(Me.lblOwnerIDValue.Text, "SELF")
                End If
            End If
            Dim MyFrm As MusterContainer
            MyFrm = Me.MdiParent
            bolFrmActivated = True
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

#End Region

#Region "Owner Operations"
    Friend Sub PopulateOwnerInfo(ByVal OwnerID As Integer)
        Dim MyFrm As MusterContainer
        Try
            UIUtilsGen.PopulateOwnerInfo(OwnerID, oOwner, Me)

            If Not tbCtrlOwner.TabPages.Contains(tbPageOwnerContactList) Then
                tbCtrlOwner.TabPages.Add(tbPageOwnerContactList)
                lblOwnerContacts.Text = "Owner Contacts"
            End If
            Select Case tbCtrlOwner.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If Me.lblOwnerIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, OwnerID, UIUtilsGen.EntityTypes.Owner)
                    End If
                Case tbPageOwnerDocuments.Name
                    UCOwnerDocuments.LoadDocumentsGrid(OwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Technical)
            End Select
            'LoadContacts(ugOwnerContacts, OwnerID, 9)
            If bolFrmActivated Then
                MyFrm = Me.MdiParent
                If lblOwnerIDValue.Text <> String.Empty Then
                    MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, UIUtilsGen.EntityTypes.Owner, "Technical", Me.Text)
                End If
            End If
            CommentsMaintenance(, , True)
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Owner Info" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            'Nothing
        End Try


    End Sub
    Private Sub ugOwnerDetailsFacilities_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugFacilityList.DoubleClick
        Try
            If Me.Cursor Is Cursors.WaitCursor Then
                Exit Sub
            End If

            If bolTabClick = False Then
                If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If
            End If
            Me.tbCntrlTechnical.SelectedIndex = 1
            nCurrentFacility = ugFacilityList.ActiveRow.Cells.Item("FacilityID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
            PopFacility(nCurrentFacility)
            lblEventID.Text = ""

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Event Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub btnOwnerComment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerComment.Click
        Try
            CommentsMaintenance(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Facility Operations"
#Region " Facility - Previous and Next Button Operations"

    'PREVIOUS AND NEXT BUTTON OPERATIONS
    ' Below contains functionality for the next and previous buttons on the "Facility" Tab

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
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
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

    Private Sub lnkLblNextFac_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextFac.LinkClicked
        Try
            If IsNumeric(lblFacilityIDValue.Text) Then
                nCurrentFacility = GetPrevNextFacility(lblFacilityIDValue.Text, True)
                PopFacility(nCurrentFacility)
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Technical")
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Next Facility" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub lnkLblPrevFacility_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevFacility.LinkClicked
        Try
            If IsNumeric(lblFacilityIDValue.Text) Then
                nCurrentFacility = GetPrevNextFacility(lblFacilityIDValue.Text, False)
                PopFacility(nCurrentFacility)
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Technical")
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Previous Facility" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

#End Region

#Region "Facility Form Events"
    Private Sub txtFacilityLongDegree_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityLongDegree.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillSingleObjectValues(oOwner.Facilities.LongitudeDegree, txtFacilityLongDegree.Text.Trim)
        Check_If_Datum_Enable()
    End Sub
    Private Sub txtFacilityLongMin_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityLongMin.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillSingleObjectValues(oOwner.Facilities.LongitudeMinutes, txtFacilityLongMin.Text.Trim)
        Check_If_Datum_Enable()
    End Sub
    Private Sub txtFacilityLongSec_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityLongSec.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillDoubleObjectValues(oOwner.Facilities.LongitudeSeconds, txtFacilityLongSec.Text.Trim)
        Check_If_Datum_Enable()
    End Sub
    Private Sub txtFacilityLatDegree_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityLatDegree.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillSingleObjectValues(oOwner.Facilities.LatitudeDegree, txtFacilityLatDegree.Text.Trim)
        Check_If_Datum_Enable()
    End Sub
    Private Sub txtFacilityLatMin_Validating(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityLatMin.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillSingleObjectValues(oOwner.Facilities.LatitudeMinutes, txtFacilityLatMin.Text.Trim)
        Check_If_Datum_Enable()
    End Sub
    Private Sub txtFacilityLatSec_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityLatSec.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillDoubleObjectValues(oOwner.Facilities.LatitudeSeconds, txtFacilityLatSec.Text.Trim)
        Check_If_Datum_Enable()
    End Sub
    Private Sub cmbFacilityDatum_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFacilityDatum.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oOwner.Facilities.Datum = UIUtilsGen.GetComboBoxValueInt(cmbFacilityDatum)
    End Sub
    Private Sub cmbFacilityMethod_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFacilityMethod.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oOwner.Facilities.Method = UIUtilsGen.GetComboBoxValueInt(cmbFacilityMethod)
    End Sub
    Private Sub cmbFacilityLocationType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFacilityLocationType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oOwner.Facilities.LocationType = UIUtilsGen.GetComboBoxValueInt(cmbFacilityLocationType)
    End Sub
    Private Sub btnFacilityCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacilityCancel.Click
        Try

            If Not oOwner.Facilities Is Nothing Then
                oOwner.Facilities.Reset()
                If oOwner.Facilities.ID > 0 Then
                    Me.PopFacility(Integer.Parse(oOwner.Facilities.ID))
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFacilitySave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacilitySave.Click
        Try
            If oOwner.Facilities.ID <= 0 Then
                oOwner.Facilities.CreatedBy = MusterContainer.AppUser.ID
            Else
                oOwner.Facilities.ModifiedBy = MusterContainer.AppUser.ID
            End If

            oOwner.Facilities.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
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
                cmdSQLCommand.CommandText = "update tblReg_AssessDate set AssessDate = " + IIf(dtPickAssess.Checked, "'" + dtPickAssess.Value.ToShortDateString + "'", "NULL") + " where FacilityID = " + lblFacilityIDValue.Text
            Else
                cmdSQLCommand.CommandText = "insert into tblReg_AssessDate values(" + lblFacilityIDValue.Text + "," + IIf(dtPickAssess.Checked, "'" + dtPickAssess.Value.ToShortDateString + "'", "NULL") + ")"
            End If
            aReader.Close()
            cmdSQLCommand.ExecuteNonQuery()
            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()
            MsgBox("Facility Information Saved Successfully!")
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub pnl_FacilityDetail_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnl_FacilityDetail.Paint
        If oFacility Is Nothing Then
            MsgBox("This owner has no facilities.")
            Me.lnkLblNextFac.Enabled = False
            Me.lnkLblPrevFacility.Enabled = False
        End If
    End Sub
    Private Sub btnFacComments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacComments.Click
        Try
            CommentsMaintenance(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Misc Operations"
    Private Sub Handlers()

    End Sub
    Private Sub FillComboBox(ByVal cmb As ComboBox, ByVal dttable As DataTable)
        Try
            cmb.DataSource = dttable
            cmb.DisplayMember = "PROPERTY_NAME"
            cmb.ValueMember = "PROPERTY_ID"
            cmb.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
    Public Sub Check_If_Datum_Enable()
        Try
            UIUtilsGen.Check_If_Datum_Enable(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Populate Operations"

    Private Sub SetToolTips()

        Dim ttPriority As New ToolTip
        Dim strTT As String

        strTT &= "PRIORITY # 1  (highest)" & vbCrLf
        strTT &= "     Vapors or free product is present in utilities, buildings, or homes." & vbCrLf
        strTT &= "     Water wells, drinking water supply lines, or surface waters are impacted or immediately threatened." & vbCrLf
        strTT &= "     Utilities are impacted or immediately threatened." & vbCrLf
        strTT &= "PRIORITY # 2" & vbCrLf
        strTT &= "     Adjacent property is impacted by BTEX/PAH groundwater or soil contamination." & vbCrLf
        strTT &= "     Free product is outside the tankbed and is not defined." & vbCrLf
        strTT &= "PRIORITY # 3" & vbCrLf
        strTT &= "     Free product is present outside of the tankbed and is defined." & vbCrLf
        strTT &= "     Free product is present in the tankbed at > 6 inches." & vbCrLf
        strTT &= "     BTEX groundwater contamination is not defined." & vbCrLf
        strTT &= "     UST system failed the precision tightness test." & vbCrLf
        strTT &= "     Commission orders." & vbCrLf
        strTT &= "PRIORITY # 4" & vbCrLf
        strTT &= "     Free product is present at < 6 inches in the tankbed." & vbCrLf
        strTT &= "     BTEX groundwater contamination is defined on site." & vbCrLf
        strTT &= "PRIORITY # 5" & vbCrLf
        strTT &= "     BTEX soil contamination is not defined." & vbCrLf
        strTT &= "     PAH groundwater contamination is not defined." & vbCrLf
        strTT &= "PRIORITY # 6" & vbCrLf
        strTT &= "     BTEX soil contamination is defined." & vbCrLf
        strTT &= "     PAH groundwater contamination is defined." & vbCrLf
        strTT &= "     PAH soil contamination is not defined." & vbCrLf
        strTT &= "PRIORITY # 7" & vbCrLf
        strTT &= "     PAH soil contamination is defined." & vbCrLf
        strTT &= "     Old sites with TPH contamination." & vbCrLf
        strTT &= "PRIORITY # 8  (lowest)" & vbCrLf
        strTT &= "     Confirmation groundwater sampling." & vbCrLf
        strTT &= "     High vapor readings in tankbed wells." & vbCrLf
        strTT &= "     Leak has not been confirmed." & vbCrLf
        strTT &= "     Monitoring leak detection records." & vbCrLf
        strTT &= "     Only tracking the site"


        ttPriority.AutoPopDelay = 15000 ' set lenth tip is visible to 15 seconds
        ttPriority.SetToolTip(lblPrioritytt, strTT)

    End Sub
    Private Sub SetPMHistory()
        Dim strTT As String
        Try
            strTT = oLustEvent.LustEventPMHistory

            ttHistory.AutoPopDelay = 15000 ' set lenth tip is visible to 15 seconds
            ttHistory.SetToolTip(lblPMHistorytt, strTT)
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Set PM History" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub FillDateobjectValues(ByRef currentObj As Object, ByVal value As String)
        If value.Length > 0 And value <> "__/__/____" Then
            currentObj = CType(value, Date)
        Else
            currentObj = "#12:00:00AM#"
        End If
    End Sub
    Friend Sub PopFacility(Optional ByVal FacilityID As Integer = 0)
        Dim MyFrm As MusterContainer
        Try
            UIUtilsGen.PopulateFacilityInfo(Me, oOwner.OwnerInfo, oOwner.Facilities, FacilityID)
            nCurrentFacility = oOwner.Facilities.ID
            nFacilityID = FacilityID
            If Not tbCntrlFacility.TabPages.Contains(tbPageOwnerContactList) Then
                tbCntrlFacility.TabPages.Add(tbPageOwnerContactList)
                lblOwnerContacts.Text = "Facility Contacts"
            End If
            Select Case tbCntrlFacility.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If Me.lblFacilityIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, FacilityID, UIUtilsGen.EntityTypes.Facility)
                    End If
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(FacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Technical)
            End Select
            'LoadContacts(ugOwnerContacts, FacilityID, 6)
            If bolFrmActivated Then
                MyFrm = Me.MdiParent
                MyFrm.FlagsChanged(Me.lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Technical", Me.Text)
            End If
            CommentsMaintenance(, , True)

            'Added by Hua Cao 10/22/2008 Issue #: [ UST-3204] Summary: Need a date field labeled " TOS Assessment Date:" added to several modules
            ' Retreive info from tblReg_AssessDate
            Me.dtPickAssess.Enabled = True
            Dim sqlStr As String
            Dim dtReturn As DataTable
            Dim dtNullDate As Date = CDate("01/01/0001")
            sqlStr = "tblReg_AssessDate where FacilityId = " + lblFacilityIDValue.Text
            dtReturn = oLustEvent.GetDataTable(sqlStr)
            If dtReturn Is Nothing Then
                dtPickAssess.Checked = False

                UIUtilsGen.SetDatePickerValue(dtPickAssess, dtNullDate)
            Else
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
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Facility" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Public Sub GetLustEventsForFacility()
        Dim drRow As DataRow
        Dim rowcount As Integer = 0
        Dim str As String = String.Empty

        Try
            strLustEventIdTags = String.Empty
            For Each drRow In oOwner.Facilities.LustEventDataset.Tables(0).Rows
                If rowcount < oOwner.Facilities.LustEventDataset.Tables(0).Rows.Count - 1 Then
                    str = ","
                Else
                    str = ""
                End If
                strLustEventIdTags += drRow("EVENT_ID").ToString + str
                rowcount += 1
            Next
            dgLUSTEvents.DataSource = oOwner.Facilities.LustEventDataset
            dgLUSTEvents.Rows.ExpandAll(True)
            dgLUSTEvents.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            dgLUSTEvents.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            If dgLUSTEvents.Rows.Count > 0 Then
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("FACILITY_ID").Hidden = True
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("EVENT_ID").Hidden = True

                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Event_Sequence").Header.Caption = "Event #"
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Event_Started").Header.Caption = "Start Date"
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Event_Ended").Header.Caption = "NFA Date"
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Confirmed_On").Header.Caption = "Confirmed"
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("ProjectManager").Header.Caption = "Project Manager"
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("EventStatus").Header.Caption = "Event Status"
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("MGPTF_Status").Header.Caption = "MGPTF Status"

                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Event_Sequence").Width = 50
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Event_Started").Width = 80
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Event_Started").Width = 80
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Priority").Width = 50
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("Confirmed_On").Width = 80
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("EventStatus").Width = 80
                dgLUSTEvents.DisplayLayout.Bands(0).Columns("MGPTF_Status").Width = 90

                If (bolPMHead Or MusterContainer.AppUser.HEAD_ADMIN) Then
                    btnTransferLUSTEvent.Visible = True
                Else
                    btnTransferLUSTEvent.Visible = False
                End If
                lblTotalNoOfLUSTEventsValue.Text = dgLUSTEvents.Rows.Count
            Else
                btnTransferLUSTEvent.Visible = False
                lblTotalNoOfLUSTEventsValue.Text = 0
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Get Lust Event Rows" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub GetLustActivitiesForEvent(ByVal nEventID As Long)
        Dim recActivityDoc As New DataSet
        Dim i As Integer
        Dim docCount As Integer
        Dim docName As String
        Dim tempStr As String
        Dim lastGWSDate As Date = CDate("1/1/1900")
        Dim receivedDate As Date
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        Try
            ugActivitiesandDocuments.DataSource = oLustEvent.LustActivityDocumentDataset(nEventID)
            ugActivitiesandDocuments.Rows.ExpandAll(True)
            recActivityDoc = oLustEvent.LustGetLastGWS(nEventID)
            docCount = recActivityDoc.Tables(0).Rows.Count() - 1
            'Get the last GWS date from the latest received date of "GWS/TRI/QUARTER/SEMI Rpt" documents
            For i = 0 To docCount
                docName = recActivityDoc.Tables(0).Rows(i).Item("Document")
                tempStr = docName.Trim()
                docName = tempStr
                If ((docName.ToUpper = "TRI RPT") Or (docName.ToUpper = "QUARTER RPT") Or (docName.ToUpper = "SEMI RPT")) Then
                    tempStr = docName
                Else
                    tempStr = docName.Remove(docName.Length() - 2, 2)
                End If
                If ((docName.ToUpper.StartsWith("GWS RPT") Or tempStr.ToUpper.EndsWith("TRI RPT") Or tempStr.ToUpper.EndsWith("QUARTER RPT") Or tempStr.ToUpper.EndsWith("SEMI RPT"))) Then
                    If Not IsDBNull(recActivityDoc.Tables(0).Rows(i).Item("Received")) Then
                        receivedDate = recActivityDoc.Tables(0).Rows(i).Item("Received")
                    End If
                    If receivedDate > lastGWSDate Then
                        lastGWSDate = receivedDate
                    End If
                End If
            Next
            'ugActivitiesandDocuments.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            lblActivitiesDocumentsCellDesc.Text = ""
            lblActivitiesDocumentsCellDesc.BackColor = SystemColors.Control
            If ugActivitiesandDocuments.Rows.Count > 0 Then
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("Activity").Width = 250
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("EVENT_ID").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("Event_Activity_ID").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("Activity_Type_ID").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("DaysSinceStart").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("Warn").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("Act_By").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("Tech_Completed_Date").Header.Caption = "Completed"
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("Closed_Date").Header.Caption = "Closed Date"
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("Start_Date").Header.Caption = "Start Date"
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("1ST_GWS_BELOW").Header.Caption = "1st GWS Below"
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("2ND_GWS_BELOW").Header.Caption = "2nd GWS Below"
                ugActivitiesandDocuments.DisplayLayout.Bands(0).Columns("ActivityDocsRelCount").Hidden = True

                ugActivitiesandDocuments.DisplayLayout.Bands(1).Columns("EVENT_ID").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(1).Columns("Event_Activity_ID").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(1).Columns("EVENT_ACTIVITY_DOCUMENT_ID").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(1).Columns("RevisionDate").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(1).Columns("COMMITMENT_ID").Hidden = True
                ugActivitiesandDocuments.DisplayLayout.Bands(1).Columns("Doc_Type").Hidden = True
                If (lastGWSDate > CDate("1/1/1900")) Then
                    If Not (oLustEvent.LastGWS = lastGWSDate) Then
                        oLustEvent.LastGWS = lastGWSDate
                        dtLastGWS.Value = lastGWSDate
                        ProcessSaveEvent()
                    End If
                End If
                ToggleActivitiesGrid()
                ToggleDocumentsGrid()
                SetActivitiesGrid()
                btnModifyActivity.Visible = True
                btnModifyDocument.Visible = True

                If bolPMHead Or cmbProjectManager.Text.ToUpper = MusterContainer.AppUser.Name.ToUpper Then
                    btnDeleteActivity.Visible = True
                    btnDeleteDocument.Visible = True
                Else
                    btnDeleteActivity.Visible = False
                    btnDeleteDocument.Visible = False
                End If
            Else
                btnDeleteActivity.Visible = False
                btnDeleteDocument.Visible = False
                btnModifyActivity.Visible = False
                btnModifyDocument.Visible = False
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Get Lust Event Activity Rows" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub GetLustRemediationForEvent(ByVal nEventID As Long)
        Try
            ugRemediationSystem.DataSource = oLustEvent.LustEventRemediationDataset(nEventID)
            ugRemediationSystem.Rows.ExpandAll(True)
            ugRemediationSystem.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugRemediationSystem.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            If ugRemediationSystem.Rows.Count > 0 Then
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("REM_SYSTEM_ID").Hidden = True
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("EVENT_ID").Hidden = True
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("Event_Activity_ID").Hidden = True
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("SYSTEM_SEQ").Hidden = True
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("SYSTEM_DEC").Hidden = True

                ugRemediationSystem.DisplayLayout.Bands(0).Columns("Start_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("Closed_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("Closed_Date").Header.Caption = "Closed Date"
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("Start_Date").Header.Caption = "Start Date"
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("System_Type").Header.Caption = "System Type"
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("Owned_Leased").Header.Caption = "Owned/Leased"
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("VACPUMP1_Size").Header.Caption = "Vac Pump1 Size"
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("VACPUMP2_Size").Header.Caption = "Vac Pump2 Size"
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("BUILDING_SIZE").Header.Caption = "Building Size"
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("Activity_Name").Header.Caption = "Activity"
                ugRemediationSystem.DisplayLayout.Bands(0).Columns("Description").Header.Caption = "Description"
                btnModifyRemediationSystem.Enabled = True
                ugRemediationSystem.Visible = True
            Else
                ugRemediationSystem.Visible = False
                btnModifyRemediationSystem.Enabled = False
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Get Lust Event Remediation Rows" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Public Sub GetLustTanksAndPipesForFacility(ByVal nEventID As Long)

        Try
            ugTankandPipes.DataSource = oOwner.Facilities.LustTankPipeDataset(nEventID)
            ugTankandPipes.Rows.ExpandAll(True)
            If ugTankandPipes.Rows.Count > 0 Then
                ugTankandPipes.DisplayLayout.Bands(0).Columns("FACILITY_ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("POSITION").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("TANK ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("DATELASTUSED").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("SPILLINSTALLED").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("OVERFILLINSTALLED").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("TANKMATDESC").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("TANKMODDESC").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("IncludedDet").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("INSTALLED").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(0).Columns("Included").Header.Caption = "Source Of Release"

                ugTankandPipes.DisplayLayout.Bands(1).Columns("FACILITY_ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("TANK ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("PIPE_ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("POSITION").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("PIPE_MAT_DESC").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("PIPE_MOD_DESC").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("PIPE_CP_TYPE").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("DATELASTUSED").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("IncludedDet").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(1).Columns("INSTALLED").Hidden = True

                ugTankandPipes.DisplayLayout.Bands(1).Columns("Filler1").Header.Caption = ""
                ugTankandPipes.DisplayLayout.Bands(1).Columns("Included").Header.Caption = "Source Of Release"

                ugTankandPipes.DisplayLayout.Bands(0).Columns("Tank Site ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(0).Columns("Capacity").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(0).Columns("Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(0).Columns("Substance").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(0).Columns("FuelType").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                ugTankandPipes.DisplayLayout.Bands(1).Columns("Pipe Site ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(1).Columns("Filler1").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(1).Columns("Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(1).Columns("Substance").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(1).Columns("FuelType").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Get Lust Event Tank and Pipe Rows" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub
#End Region

#End Region

#Region "Lust Event Operations"

#Region "Lust Event Form Events"


    Private Sub txtFacilityAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.DoubleClick
        Try

            Dim addressForm As Address
            If (oOwner.Facilities.ID = 0 Or oOwner.Facilities.ID <> nFacilityID) And nFacilityID > 0 Then
                UIUtilsGen.PopulateFacilityInfo(Me, oOwner.OwnerInfo, oOwner.Facilities, nFacilityID)
                '  UIUtilsGen.PopulateFacilityInfo(Me, oOwner.OwnerInfo, oOwner.Facilities, strFacilityIdTags)
            End If

            Address.EditAddress(addressForm, oOwner.Facilities.ID, oOwner.Facilities.FacilityAddresses, "Facility", UIUtilsGen.ModuleID.Technical, txtFacilityAddress, UIUtilsGen.EntityTypes.Facility, True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtOwnerAddress_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.DoubleClick
        Try

            Dim addressForm As Address
            Address.EditAddress(addressForm, oOwner.ID, oOwner.Addresses, "Owner", UIUtilsGen.ModuleID.Technical, txtOwnerAddress, UIUtilsGen.EntityTypes.Owner)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub pnlActivitiesDocumentsDetails_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlActivitiesDocumentsDetails.Resize
        ugActivitiesandDocuments.Width = pnlActivitiesDocumentsDetails.Width - 50

    End Sub

    Private Sub pnlRemediationSystemsDetails_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlRemediationSystemsDetails.Resize
        ugRemediationSystem.Width = pnlRemediationSystemsDetails.Width - 50
    End Sub

    Private Sub Technical_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
    End Sub

    Private Sub Technical_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If oLustEvent.IsDirty And btnSaveLUSTEvent.Enabled = True And btnSaveLUSTEvent.Visible = True Then

            If _container.DirtyIgnored = -1 OrElse (_container.DirtyIgnored = MsgBoxResult.Yes AndAlso MsgBox("Do you wish to save changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
                ProcessSaveEvent(True)
            End If

        End If

    End Sub

    Private Sub btnSaveLUSTEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveLUSTEvent.Click
        ProcessSaveEvent()
    End Sub

    Private Sub btnTransferLUSTEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTransferLUSTEvent.Click
        Dim frmXFerEvent As New TransferEvent
        Try
            If nCurrentFacility <> 0 Then
                frmXFerEvent.CallingForm = Me
                frmXFerEvent.FacilityID = nCurrentFacility
                frmXFerEvent.FacilityName = oOwner.Facilities.Name

                frmXFerEvent.ShowDialog()

                GetLustEventsForFacility()
                UIUtilsGen.PopulateOwnerFacilities(oOwner, Me, Integer.Parse(Me.lblOwnerIDValue.Text))
            Else
                MsgBox("No Facility selected.")
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Event Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmXFerEvent = Nothing
        End Try
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            'If oOwner.colIsDirty() Then
            '    Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
            '    If Results = MsgBoxResult.Yes Then
            '        Dim success As Boolean = False
            '        If oOwner.ID <= 0 Then
            '            oOwner.CreatedBy = MusterContainer.AppUser.ID
            '        Else
            '            oOwner.ModifiedBy = MusterContainer.AppUser.ID
            '        End If
            '        success = oOwner.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            '        If Not UIUtilsGen.HasRights(returnVal) Then
            '            Exit Sub
            '        End If

            '        If Not success Then
            '            e.Cancel = True
            '            Exit Sub
            '        End If
            '    ElseIf Results = MsgBoxResult.Cancel Then
            '        e.Cancel = True
            '        Exit Sub
            '    End If
            'End If
            'if any other forms are using the owner, leave alone. else remove from collection
            UIUtilsGen.RemoveOwner(oOwner, Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ugTankandPipes_AfterCellUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTankandPipes.AfterCellUpdate
        ProcessTankAndPipe()
    End Sub

    Private Function EligibityIsChecked() As Boolean
        Dim bolReturn As Boolean = False

        bolReturn = Me.chkCommissionNo.Checked Or _
                    Me.chkCommissionYes.Checked Or _
                    Me.chkOPCHeadNo.Checked Or _
                    Me.chkOPCHeadYes.Checked Or _
                    Me.chkPMHeadUndecided.Checked Or _
                    Me.chkPMHeadYes.Checked Or _
                    Me.chkUSTChiefNo.Checked Or _
                    Me.chkUSTChiefUndecided.Checked Or _
                    Me.chkUSTChiefYes.Checked

        EligibityIsChecked = bolReturn

    End Function

    Private Function EligibityIsComplete() As Boolean
        Dim bolReturn As Boolean = False

        bolReturn = Me.chkCommissionNo.Checked Or _
                    Me.chkCommissionYes.Checked Or _
                    Me.chkOPCHeadNo.Checked Or _
                    Me.chkOPCHeadYes.Checked Or _
                    Me.chkPMHeadYes.Checked Or _
                    Me.chkUSTChiefNo.Checked Or _
                    Me.chkUSTChiefYes.Checked

        EligibityIsComplete = bolReturn

    End Function
    Private Sub btnDeleteActivity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteActivity.Click
        Dim oLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
        Dim bolResponse As MsgBoxResult
        Try
            oLustActivity.Retrieve(ugActivitiesandDocuments.ActiveRow.Cells("Event_Activity_ID").Value)

            If ugActivitiesandDocuments.ActiveRow.Band.Index = 0 Then
                If ugActivitiesandDocuments.ActiveRow.ChildBands(0).Rows.Count = 0 And Not (oLustActivity.RemSystemID > 0) Then
                    bolResponse = MsgBox("Are you sure you want to delete this Activity?", MsgBoxStyle.YesNo, "Delete Activity Conformation")
                    If bolResponse = MsgBoxResult.Yes Then
                        oLustActivity.Deleted = True
                        oLustActivity.ModifiedBy = MusterContainer.AppUser.ID
                        oLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        GetLustActivitiesForEvent(oLustEvent.ID)
                        GetLustRemediationForEvent(oLustEvent.ID)
                    End If
                Else
                    MsgBox("You cannot delete an activity with documents or remediation systems.")
                End If
            Else
                MsgBox("You must select an activity to delete.")
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Delete Activity" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnDeleteDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteDocument.Click
        Dim oLustDocument As New MUSTER.BusinessLogic.pLustEventDocument
        Dim oCalendar As New MUSTER.BusinessLogic.pCalendar
        Dim dtTest As Date
        Dim bolResponse As MsgBoxResult
        Dim iDocumentID As Int64
        Dim MyFrm As MusterContainer

        Try
            If ugActivitiesandDocuments.ActiveRow.Band.Index = 1 Then
                If ugActivitiesandDocuments.ActiveRow.Cells("Received").Value Is DBNull.Value Then
                    If ugActivitiesandDocuments.ActiveRow.Cells("COMMITMENT_ID").Value = 0 Then
                        bolResponse = MsgBox("Are you sure you want to delete this Document?", MsgBoxStyle.YesNo, "Delete Document Conformation")
                        If bolResponse = MsgBoxResult.Yes Then
                            oLustDocument.Retrieve(ugActivitiesandDocuments.ActiveRow.Cells("EVENT_ACTIVITY_DOCUMENT_ID").Value)
                            oLustDocument.Deleted = True
                            iDocumentID = oLustDocument.ID
                            oLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                            oLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If
                            GetLustActivitiesForEvent(oLustEvent.ID)
                            GetLustRemediationForEvent(oLustEvent.ID)

                            MyFrm = MdiParent
                            MyFrm.RefreshCalendarInfo()
                            MyFrm.LoadDueToMeCalendar()
                            MyFrm.LoadToDoCalendar()
                        End If
                    Else
                        MsgBox("You cannot delete a document with associated commitment.")
                    End If
                Else
                    MsgBox("You cannot delete a document with a received date.")
                End If
            Else
                MsgBox("You must select a document to delete.")
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Delete Document" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub


    Private Sub btnSendtoPM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendtoPM.Click
        Try
            oLustEvent.CalendarEntries(PMHead_UserID, "Review Lust Event Checklist for ID: " & oLustEvent.FacilityID & " Event " & oLustEvent.EVENTSEQUENCE, True, False, String.Empty, UserID, Now.Date, Now.Date, oLustEvent.ID)
            MsgBox("Checklist sent successfully to PM-Head")
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Send to PM" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub


    Private Sub btnAddLUSTEvent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddLUSTEvent.Click
        Try
            lblEventID.Text = "Add Event"
            lblEventIDValue.Visible = False
            lblEventCountValue.Visible = False
            nLastEventID = -1
            oLustEvent.Retrieve(0)

            SetupAddLustEventForm()
            Me.Text = "Technical - Add Lust Event (" & oOwner.Facilities.ID & ")"

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Add Lust Event" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub dgLUSTEvents_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgLUSTEvents.DoubleClick
        Dim MyFrm As MusterContainer

        If Me.Cursor Is Cursors.WaitCursor Then
            Exit Sub
        End If
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            Me.tbCntrlTechnical.SelectedTab = Me.tbPageLUSTEvent
            nSavedEventID = 0
            SetupModifyLustEventForm()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Lust Event Double Click" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try


    End Sub

    Private Sub btnEvtTankToggle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEvtTankToggle.Click

        EventTankToggle()

    End Sub

    Private Sub btnEvtTankCollapse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEvtTankCollapse.Click
        If btnEvtTankCollapse.Text = "Collapse" Then
            ugTankandPipes.Rows.CollapseAll(True)
            btnEvtTankCollapse.Text = "Expand"
        Else
            ugTankandPipes.Rows.ExpandAll(True)
            btnEvtTankCollapse.Text = "Collapse"
        End If
    End Sub
    Private Sub chkSoil_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSoil.CheckedChanged
        If chkSoil.Checked Then
            chkSoilBTEX.Enabled = True
            chkSoilPAH.Enabled = True
            chkSoilTPH.Enabled = True
        Else
            chkSoilBTEX.Enabled = False
            chkSoilPAH.Enabled = False
            chkSoilTPH.Enabled = False

            chkSoilBTEX.Checked = False
            chkSoilPAH.Checked = False
            chkSoilTPH.Checked = False
        End If
        If bolLoading Then Exit Sub
        oLustEvent.TOCSOIL = chkSoil.Checked
    End Sub
    Private Sub chkGroundWater_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGroundWater.CheckedChanged
        If chkGroundWater.Checked Then
            chkGroundWaterBTEX.Enabled = True
            chkGroundWaterPAH.Enabled = True
            chkGroundWaterTPH.Enabled = True
        Else

            chkGroundWaterBTEX.Enabled = False
            chkGroundWaterPAH.Enabled = False
            chkGroundWaterTPH.Enabled = False

            chkGroundWaterBTEX.Checked = False
            chkGroundWaterPAH.Checked = False
            chkGroundWaterTPH.Checked = False
        End If
        If bolLoading Then Exit Sub
        oLustEvent.TOCGROUNDWATER = chkGroundWater.Checked
    End Sub
    Private Sub chkToCFreeProduct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkToCFreeProduct.CheckedChanged
        If chkToCFreeProduct.Checked Then
            chkFreeProductDiesel.Enabled = True
            chkFreeProductGasoline.Enabled = True
            chkFreeProductKerosene.Enabled = True
            chkFreeProductUnKnown.Enabled = True
            chkFreeProductWasteOil.Enabled = True
        Else
            chkFreeProductDiesel.Enabled = False
            chkFreeProductGasoline.Enabled = False
            chkFreeProductKerosene.Enabled = False
            chkFreeProductUnKnown.Enabled = False
            chkFreeProductWasteOil.Enabled = False

            chkFreeProductDiesel.Checked = False
            chkFreeProductGasoline.Checked = False
            chkFreeProductKerosene.Checked = False
            chkFreeProductUnKnown.Checked = False
            chkFreeProductWasteOil.Checked = False
        End If
        If bolLoading Then Exit Sub
        oLustEvent.FREEPRODUCT = chkToCFreeProduct.Checked
    End Sub
    Private Sub chkToCVapor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkToCVapor.CheckedChanged
        If chkToCVapor.Checked Then
            chkVaporBTEX.Enabled = True
            chkVaporPAH.Enabled = True
        Else
            chkVaporBTEX.Enabled = False
            chkVaporPAH.Enabled = False

            chkVaporBTEX.Checked = False
            chkVaporPAH.Checked = False
        End If
        If bolLoading Then Exit Sub
        oLustEvent.TOCVAPOR = chkToCVapor.Checked
    End Sub

    'Private Sub chkHowDiscovered_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFacilityLeakDetection.CheckedChanged, chkGWWell.CheckedChanged, chkSurfaceSheen.CheckedChanged, chkGWContamination.CheckedChanged, chkVapors.CheckedChanged, chkFreeProduct.CheckedChanged, chkSoilContamination.CheckedChanged, chkFailedPTT.CheckedChanged, chkInventoryShortage.CheckedChanged, chkTankClosure.CheckedChanged, chkInspection.CheckedChanged


    '    If sender.checked Then
    '        If sender.name <> chkFacilityLeakDetection.Name Then
    '            chkFacilityLeakDetection.Checked = False
    '        End If

    '        If sender.name <> chkGWWell.Name Then
    '            chkGWWell.Checked = False
    '        End If

    '        If sender.name <> chkSurfaceSheen.Name Then
    '            chkSurfaceSheen.Checked = False
    '        End If

    '        If sender.name <> chkGWContamination.Name Then
    '            chkGWContamination.Checked = False
    '        End If

    '        If sender.name <> chkVapors.Name Then
    '            chkVapors.Checked = False
    '        End If

    '        If sender.name <> chkFreeProduct.Name Then
    '            chkFreeProduct.Checked = False
    '        End If

    '        If sender.name <> chkSoilContamination.Name Then
    '            chkSoilContamination.Checked = False
    '        End If

    '        If sender.name <> chkFailedPTT.Name Then
    '            chkFailedPTT.Checked = False
    '        End If

    '        If sender.name <> chkInventoryShortage.Name Then
    '            chkInventoryShortage.Checked = False
    '        End If

    '        If sender.name <> chkTankClosure.Name Then
    '            chkTankClosure.Checked = False
    '        End If

    '        If sender.name <> chkInspection.Name Then
    '            chkInspection.Checked = False
    '        End If
    '        If bolLoading = False Then
    '            oLustEvent.HowDiscoveredID = sender.tag
    '        End If


    '    End If

    '    If bolLoading Then Exit Sub
    '    If chkFacilityLeakDetection.Checked = False And chkGWWell.Checked = False And chkSurfaceSheen.Checked = False And chkGWContamination.Checked = False _
    '        And chkVapors.Checked = False And chkFreeProduct.Checked = False And chkSoilContamination.Checked = False And chkFailedPTT.Checked = False _
    '        And chkInventoryShortage.Checked = False And chkTankClosure.Checked = False And chkInspection.Checked = False Then
    '        oLustEvent.HowDiscoveredID = 0
    '    End If

    'End Sub

    Private Sub chkFacilityLeakDetection_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFacilityLeakDetection.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscFacLD = chkFacilityLeakDetection.Checked
    End Sub
    Private Sub chkSurfaceSheen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSurfaceSheen.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscSurfaceSheen = chkSurfaceSheen.Checked
    End Sub
    Private Sub chkGWWell_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGWWell.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscGWWell = chkGWWell.Checked
    End Sub
    Private Sub chkGWContamination_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGWContamination.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscGWContamination = chkGWContamination.Checked
    End Sub
    Private Sub chkVapors_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkVapors.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscVapors = chkVapors.Checked
    End Sub
    Private Sub chkFreeProduct_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFreeProduct.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscFreeProduct = chkFreeProduct.Checked
    End Sub
    Private Sub chkSoilContamination_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSoilContamination.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscSoilContamination = chkSoilContamination.Checked
    End Sub
    Private Sub chkFailedPTT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFailedPTT.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscFailedPTT = chkFailedPTT.Checked
    End Sub
    Private Sub chkInventoryShortage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkInventoryShortage.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscInventoryShortage = chkInventoryShortage.Checked
    End Sub
    Private Sub chkTankClosure_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTankClosure.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscTankClosure = chkTankClosure.Checked
    End Sub
    Private Sub chkInspection_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkInspection.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.HowDiscInspection = chkInspection.Checked
    End Sub

    Private Sub dtPMHeadOn_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPMHeadOn.ValueChanged
        If bolLoading Then Exit Sub
        oLustEvent.PM_HEAD_DATE = dtPMHeadOn.Value.Date
    End Sub

    Private Sub dtUSTChiefOn_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtUSTChiefOn.ValueChanged
        If bolLoading Then Exit Sub
        oLustEvent.UST_CHIEF_DATE = dtUSTChiefOn.Value.Date
    End Sub

    Private Sub dtOPCHeadOn_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtOPCHeadOn.ValueChanged
        If bolLoading Then Exit Sub
        oLustEvent.OPC_HEAD_DATE = dtOPCHeadOn.Value.Date
    End Sub

    Private Sub dtCommissionOn_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtCommissionOn.ValueChanged
        If bolLoading Then Exit Sub
        oLustEvent.COMMISSION_DATE = dtCommissionOn.Value.Date
    End Sub

    Private Sub txtPMHeadBy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPMHeadBy.TextChanged
        If bolLoading Then Exit Sub
        oLustEvent.PM_HEAD_BY = txtPMHeadBy.Text
    End Sub

    Private Sub txtUSTChiefBy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUSTChiefBy.TextChanged
        If bolLoading Then Exit Sub
        oLustEvent.UST_CHIEF_BY = txtUSTChiefBy.Text
    End Sub

    Private Sub txtOPCHeadBy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOPCHeadBy.TextChanged
        If bolLoading Then Exit Sub
        oLustEvent.OPC_HEAD_BY = txtOPCHeadBy.Text
    End Sub

    Private Sub txtCommissionBy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCommissionBy.TextChanged
        If bolLoading Then Exit Sub
        oLustEvent.COMMISSION_BY = txtCommissionBy.Text
    End Sub
    Private Sub cmbProjectManager_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProjectManager.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.PM = cmbProjectManager.SelectedValue
    End Sub
    Private Sub cmbEventStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbEventStatus.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.EventStatus = cmbEventStatus.SelectedValue
    End Sub
    Private Sub cmbMGPTFStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMGPTFStatus.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.MGPTFStatus = cmbMGPTFStatus.SelectedValue
    End Sub
    Private Sub cmbReleaseStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbReleaseStatus.SelectedIndexChanged
        Dim dttest As Date
        If bolLoading Then Exit Sub
        oLustEvent.ReleaseStatus = cmbReleaseStatus.SelectedValue
        If cmbReleaseStatus.SelectedValue = 623 Then
            cmbSuspectedSource.Enabled = False
        Else
            cmbSuspectedSource.Enabled = True
        End If
    End Sub
    Private Sub cmbPriority_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPriority.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.Priority = cmbPriority.SelectedValue
    End Sub
    Private Sub cmbSuspectedSource_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSuspectedSource.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.SuspectedSource = cmbSuspectedSource.SelectedValue
    End Sub
    Private Sub dtStartDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtStartDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtStartDate)
        Me.FillDateobjectValues(oLustEvent.Started, dtStartDate.Text)
    End Sub
    Private Sub dtDateofReport_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtDateofReport.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtDateofReport)
        Me.FillDateobjectValues(oLustEvent.ReportDate, dtDateofReport.Text)
    End Sub
    Private Sub dtCompAssDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtCompAssDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtCompAssDate)
        Me.FillDateobjectValues(oLustEvent.CompAssDate, dtCompAssDate.Text)
    End Sub
    Private Sub dtLastGWS_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtLastGWS.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtLastGWS)
        Me.FillDateobjectValues(oLustEvent.LastGWS, dtLastGWS.Text)
    End Sub
    Private Sub dtLastLDR_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtLastLDR.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtLastLDR)
        Me.FillDateobjectValues(oLustEvent.LastLDR, dtLastLDR.Text)
    End Sub
    Private Sub dtLastPTT_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtLastPTT.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtLastPTT)
        Me.FillDateobjectValues(oLustEvent.LastPTT, dtLastPTT.Text)
    End Sub
    Private Sub txtRelatedSites_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRelatedSites.TextChanged
        If bolLoading Then Exit Sub
        oLustEvent.RelatedSites = txtRelatedSites.Text
    End Sub
    Private Sub chkSoilBTEX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSoilBTEX.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.SOILBTEX = chkSoilBTEX.Checked
    End Sub
    Private Sub chkSoilPAH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSoilPAH.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.SOILPAH = chkSoilPAH.Checked
    End Sub
    Private Sub chkSoilTPH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSoilTPH.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.SOILTPH = chkSoilTPH.Checked
    End Sub
    Private Sub chkGroundWaterBTEX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGroundWaterBTEX.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.GWBTEX = chkGroundWaterBTEX.Checked
    End Sub
    Private Sub chkGroundWaterPAH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGroundWaterPAH.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.GWPAH = chkGroundWaterPAH.Checked
    End Sub
    Private Sub chkGroundWaterTPH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGroundWaterTPH.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.GWTPH = chkGroundWaterTPH.Checked
    End Sub
    Private Sub chkFreeProductGasoline_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFreeProductGasoline.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.FPGASOLINE = chkFreeProductGasoline.Checked
    End Sub
    Private Sub chkFreeProductDiesel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFreeProductDiesel.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.FPDIESEL = chkFreeProductDiesel.Checked
    End Sub
    Private Sub chkFreeProductKerosene_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFreeProductKerosene.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.FPKEROSENE = chkFreeProductKerosene.Checked
    End Sub
    Private Sub chkFreeProductWasteOil_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFreeProductWasteOil.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.FPWASTEOIL = chkFreeProductWasteOil.Checked
    End Sub
    Private Sub chkFreeProductUnKnown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFreeProductUnKnown.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.FPUNKNOWN = chkFreeProductUnKnown.Checked
    End Sub
    Private Sub chkVaporBTEX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkVaporBTEX.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.VAPORBTEX = chkVaporBTEX.Checked
    End Sub
    Private Sub chkVaporPAH_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkVaporPAH.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.VAPORPAH = chkVaporPAH.Checked
    End Sub
    Private Sub cmbLocation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLocation.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.Location = cmbLocation.SelectedValue
        TestForConfirmedSpill()
    End Sub
    Private Sub cmbIdentifiedBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbIdentifiedBy.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.IDENTIFIEDBY = cmbIdentifiedBy.SelectedValue
    End Sub
    Private Sub cmbExtent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtent.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.Extent = cmbExtent.SelectedValue
        TestForConfirmedSpill()
    End Sub
    Private Sub cmbCause_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCause.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustEvent.Cause = cmbCause.SelectedValue
    End Sub

    Private Sub chkQuestion1Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion1Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion1Yes.Checked Then
            chkQuestion1No.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False

    End Sub
    Private Sub chkQuestion1No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion1No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion1No.Checked Then
            chkQuestion1Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion2Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion2Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion2Yes.Checked Then
            chkQuestion2No.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion2No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion2No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion2No.Checked Then
            chkQuestion2Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion5Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion5Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion5Yes.Checked Then
            chkQuestion5No.Checked = False
            chkQuestion5NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion5No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion5No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion5No.Checked Then
            chkQuestion5Yes.Checked = False
            chkQuestion5NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
            Me.bolMGPTFCheckList = True
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion5NA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion5NA.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion5NA.Checked Then
            chkQuestion5No.Checked = False
            chkQuestion5Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion6Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion6Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion6Yes.Checked Then
            chkQuestion6No.Checked = False
            chkQuestion6NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion6No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion6No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion6No.Checked Then
            chkQuestion6Yes.Checked = False
            chkQuestion6NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion6NA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion6NA.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion6NA.Checked Then
            chkQuestion6No.Checked = False
            chkQuestion6Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion8Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion8Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion8Yes.Checked Then
            chkQuestion8No.Checked = False
            chkQuestion8NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion8No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion8No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion8No.Checked Then
            chkQuestion8Yes.Checked = False
            chkQuestion8NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion8NA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion8NA.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion8NA.Checked Then
            chkQuestion8No.Checked = False
            chkQuestion8Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion10Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion10Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion10Yes.Checked Then
            chkQuestion10No.Checked = False
            chkQuestion10NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion10No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion10No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion10No.Checked Then
            chkQuestion10Yes.Checked = False
            chkQuestion10NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion10NA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion10NA.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion10NA.Checked Then
            chkQuestion10No.Checked = False
            chkQuestion10Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion14Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion14Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion14Yes.Checked Then
            chkQuestion14No.Checked = False
            chkQuestion14NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion14No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion14No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion14No.Checked Then
            chkQuestion14Yes.Checked = False
            chkQuestion14NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion14NA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion14NA.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion14NA.Checked Then
            chkQuestion14No.Checked = False
            chkQuestion14Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion15Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion15Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion15Yes.Checked Then
            chkQuestion15No.Checked = False
            chkQuestion15NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion15No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion15No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion15No.Checked Then
            chkQuestion15Yes.Checked = False
            chkQuestion15NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion15NA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion15NA.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion15NA.Checked Then
            chkQuestion15No.Checked = False
            chkQuestion15Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion16Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion16Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion16Yes.Checked Then
            chkQuestion16No.Checked = False
            chkQuestion16NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion16No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion16No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion16No.Checked Then
            chkQuestion16Yes.Checked = False
            chkQuestion16NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion16NA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion16NA.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion16NA.Checked Then
            chkQuestion16No.Checked = False
            chkQuestion16Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion17Yes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion17Yes.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion17Yes.Checked Then
            chkQuestion17No.Checked = False
            chkQuestion17NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion17No_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion17No.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion17No.Checked Then
            chkQuestion17Yes.Checked = False
            chkQuestion17NA.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkQuestion17NA_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuestion17NA.CheckedChanged
        If bolLoading Then Exit Sub
        If bolProcessingCheck Then Exit Sub
        bolProcessingCheck = True
        If chkQuestion17NA.Checked Then
            chkQuestion17No.Checked = False
            chkQuestion17Yes.Checked = False
        End If
        If lblEventID.Text <> "Add Event" Then
            Me.bolMGPTFCheckList = True
            'If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
            '    EvaluateTFChecklist(True)
            'Else
            '    EvaluateTFChecklist(False)
            'End If
        Else
            EvaluateTFChecklist(False)
        End If
        bolProcessingCheck = False
    End Sub
    Private Sub chkPMHeadYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPMHeadYes.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub
        Try
            If chkPMHeadYes.Checked Then
                chkPMHeadUndecided.Checked = False
                pnlFundsEligibilityQuestions.Enabled = False

                oLustEvent.PM_HEAD_ASSESS = 1
                cmbMGPTFStatus.SelectedValue = 618 'STFS
                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtPMHeadOn, Now.Date)
                    txtPMHeadBy.Text = UserID
                End If
            Else
                If chkPMHeadUndecided.Checked = False And chkPMHeadYes.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtPMHeadOn, tmpdate)
                    txtPMHeadBy.Text = String.Empty
                    oLustEvent.PM_HEAD_ASSESS = 0
                End If

                If chkPMHeadYes.Checked = False And chkUSTChiefYes.Checked = False And chkOPCHeadYes.Checked = False And chkCommissionYes.Checked = False Then
                    pnlFundsEligibilityQuestions.Enabled = True
                End If
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing PM Head Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try



    End Sub
    Private Sub chkPMHeadUndecided_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPMHeadUndecided.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub
        Try

            If chkPMHeadUndecided.Checked Then
                chkPMHeadYes.Checked = False
                oLustEvent.PM_HEAD_ASSESS = 3
                pnlFundsEligibilityQuestions.Enabled = True
                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtPMHeadOn, Now.Date)
                    txtPMHeadBy.Text = UserID
                End If
            Else
                If Me.chkPMHeadUndecided.Checked = False And Me.chkPMHeadYes.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtPMHeadOn, tmpdate)
                    txtPMHeadBy.Text = String.Empty
                    oLustEvent.PM_HEAD_ASSESS = 0
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing PM Head Undecided Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub chkUSTChiefYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUSTChiefYes.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub
        Try

            If chkUSTChiefYes.Checked Then
                chkUSTChiefNo.Checked = False
                chkUSTChiefUndecided.Checked = False
                oLustEvent.UST_CHIEF_ASSESS = 1
                pnlFundsEligibilityQuestions.Enabled = False

                cmbMGPTFStatus.SelectedValue = 618 'STFS
                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtUSTChiefOn, Now.Date)
                    txtUSTChiefBy.Text = UserID
                End If
            Else
                If Me.chkUSTChiefYes.Checked = False And Me.chkUSTChiefNo.Checked = False And chkUSTChiefUndecided.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtUSTChiefOn, tmpdate)
                    txtUSTChiefBy.Text = String.Empty
                    oLustEvent.UST_CHIEF_ASSESS = 0
                End If
                If chkPMHeadYes.Checked = False And chkUSTChiefYes.Checked = False And chkOPCHeadYes.Checked = False And chkCommissionYes.Checked = False Then
                    pnlFundsEligibilityQuestions.Enabled = True
                End If
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing UST Chief Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub chkUSTChiefNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUSTChiefNo.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub
        Try


            If chkUSTChiefNo.Checked Then
                chkUSTChiefYes.Checked = False
                chkUSTChiefUndecided.Checked = False
                oLustEvent.UST_CHIEF_ASSESS = 2

                cmbMGPTFStatus.SelectedValue = 620 'NTFE
                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtUSTChiefOn, Now.Date)
                    txtUSTChiefBy.Text = UserID
                End If
            Else
                If Me.chkUSTChiefYes.Checked = False And Me.chkUSTChiefNo.Checked = False And chkUSTChiefUndecided.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtUSTChiefOn, tmpdate)
                    txtUSTChiefBy.Text = String.Empty
                    oLustEvent.UST_CHIEF_ASSESS = 0
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing UST Chief Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub chkUSTChiefUndecided_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUSTChiefUndecided.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub
        Try

            If chkUSTChiefUndecided.Checked Then
                chkUSTChiefNo.Checked = False
                chkUSTChiefYes.Checked = False
                oLustEvent.UST_CHIEF_ASSESS = 3

                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtUSTChiefOn, Now.Date)
                    txtUSTChiefBy.Text = UserID
                End If
            Else
                If Me.chkUSTChiefYes.Checked = False And Me.chkUSTChiefNo.Checked = False And chkUSTChiefUndecided.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtUSTChiefOn, tmpdate)
                    txtUSTChiefBy.Text = String.Empty
                    oLustEvent.UST_CHIEF_ASSESS = 0
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing UST Chief Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub chkOPCHeadYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOPCHeadYes.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub
        Try

            If chkOPCHeadYes.Checked Then
                chkOPCHeadNo.Checked = False
                oLustEvent.OPC_HEAD_ASSESS = 1
                pnlFundsEligibilityQuestions.Enabled = False

                cmbMGPTFStatus.SelectedValue = 618 'STFS
                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtOPCHeadOn, Now.Date)
                    txtOPCHeadBy.Text = UserID
                End If
            Else
                If Me.chkOPCHeadYes.Checked = False And Me.chkOPCHeadNo.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtOPCHeadOn, tmpdate)
                    txtOPCHeadBy.Text = String.Empty
                    oLustEvent.OPC_HEAD_ASSESS = 0
                End If
                If chkPMHeadYes.Checked = False And chkUSTChiefYes.Checked = False And chkOPCHeadYes.Checked = False And chkCommissionYes.Checked = False Then
                    pnlFundsEligibilityQuestions.Enabled = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing OPC Head Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub
    Private Sub chkOPCHeadNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOPCHeadNo.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub
        Try


            If chkOPCHeadNo.Checked Then
                chkOPCHeadYes.Checked = False
                oLustEvent.OPC_HEAD_ASSESS = 2

                cmbMGPTFStatus.SelectedValue = 620 'NTFE
                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtOPCHeadOn, Now.Date)
                    txtOPCHeadBy.Text = UserID
                End If
            Else
                If Me.chkOPCHeadYes.Checked = False And Me.chkOPCHeadNo.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtOPCHeadOn, tmpdate)
                    txtOPCHeadBy.Text = String.Empty
                    oLustEvent.OPC_HEAD_ASSESS = 0
                End If
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing OPC Head Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkCommissionYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCommissionYes.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub
        Try


            If chkCommissionYes.Checked Then
                chkCommissionNo.Checked = False
                oLustEvent.COMMISSION_ASSESS = 1
                pnlFundsEligibilityQuestions.Enabled = False

                cmbMGPTFStatus.SelectedValue = 618 'STFS
                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtCommissionOn, Now.Date)
                    txtCommissionBy.Text = UserID
                End If
            Else
                If chkCommissionYes.Checked = False And chkCommissionNo.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtCommissionOn, tmpdate)
                    txtCommissionBy.Text = String.Empty
                    oLustEvent.COMMISSION_ASSESS = 0
                End If
                If chkPMHeadYes.Checked = False And chkUSTChiefYes.Checked = False And chkOPCHeadYes.Checked = False And chkCommissionYes.Checked = False Then
                    pnlFundsEligibilityQuestions.Enabled = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing Commission Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub chkCommissionNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCommissionNo.CheckedChanged
        Dim tmpdate As Date
        If bolLoading Then Exit Sub

        Try


            If chkCommissionNo.Checked Then
                chkCommissionYes.Checked = False
                oLustEvent.COMMISSION_ASSESS = 2

                cmbMGPTFStatus.SelectedValue = 620 'NTFE
                If Not bolLoading Then
                    UIUtilsGen.SetDatePickerValue(dtCommissionOn, Now.Date)
                    txtCommissionBy.Text = UserID
                End If
            Else
                If chkCommissionYes.Checked = False And chkCommissionNo.Checked = False Then
                    UIUtilsGen.SetDatePickerValue(dtCommissionOn, tmpdate)
                    txtCommissionBy.Text = String.Empty
                    oLustEvent.COMMISSION_ASSESS = 0
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Testing Commission Checked" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub
    Private Sub dtConfirmedOn_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtConfirmedOn.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtConfirmedOn)
        FillDateobjectValues(oLustEvent.Confirmed, dtConfirmedOn.Text)
        TestForConfirmedSpill()
    End Sub
    Private Sub chkShowallActivities_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowallActivities.CheckedChanged
        ToggleActivitiesGrid()
    End Sub
    Private Sub chkShowallDocuments_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowallDocuments.CheckedChanged
        ToggleDocumentsGrid()
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try


            oLustEvent.Reset()
            oLustEvent.IsDirty = False

            If lblEventID.Text = "Add Event" Then
                SetupAddLustEventForm()
                nLastEventID = -1
            Else
                SetupModifyLustEventForm(True)
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Cancel" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub btnDeleteLUSTEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteLUSTEvent.Click
        ' If event is 'Confirmed' or a Closure event, prevent the delete.
        ' Note: Closure events have release status of 'Confirmed' as well, 
        '       so no need for redundant check.  Only testing for 'Confirmed'
        Dim MyFrm As MusterContainer

        Try
            MyFrm = MdiParent
            If MsgBox("Are you sure you want to delete this Event?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                If cmbReleaseStatus.Text = "Confirmed" Then
                    MsgBox("Event cannot be deleted:  Confirmed Release")
                    Exit Sub
                End If
                oLustEvent.Deleted = True
                oLustEvent.ModifiedBy = MusterContainer.AppUser.ID
                oLustEvent.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                GetLustEventsForFacility()
                UIUtilsGen.PopulateOwnerFacilities(oOwner, Me, Integer.Parse(Me.lblOwnerIDValue.Text))
                tbCntrlTechnical.SelectedIndex = 1
                MsgBox("Technical Event Deleted")

                MyFrm.RefreshCalendarInfo()
                MyFrm.LoadDueToMeCalendar()
                MyFrm.LoadToDoCalendar()
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Delete Lust Event" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub chkforHeadofOPC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkforHeadofOPC.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.FOR_OPC_HEAD = chkforHeadofOPC.Checked
    End Sub
    Private Sub chkforCommission_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkforCommission.CheckedChanged
        If bolLoading Then Exit Sub
        oLustEvent.FOR_COMMISSION = chkforCommission.Checked
    End Sub
    Private Sub btnAddActivity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddActivity.Click
        Dim frmAddActivity As New Activity(oLustEvent)
        Dim MyFrm As MusterContainer

        Try
            MyFrm = MdiParent

            frmAddActivity.CallingForm = Me
            frmAddActivity.Mode = 0 ' Add
            frmAddActivity.EventActivityID = 0
            frmAddActivity.TFStatus = oLustEvent.MGPTFStatus
            frmAddActivity.ShowDialog()

            GetLustActivitiesForEvent(oLustEvent.ID)
            GetLustRemediationForEvent(oLustEvent.ID)

            MyFrm.RefreshCalendarInfo()
            MyFrm.LoadDueToMeCalendar()
            MyFrm.LoadToDoCalendar()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Add Activity" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmAddActivity = Nothing
        End Try
    End Sub
    Private Sub btnModifyActivity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifyActivity.Click

        ModifyActivity()

    End Sub
    Private Sub btnAddDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddDocument.Click
        'If Activity = REM Dual Phase or REM AS/SVE, OR REM Pump and Treat then 
        ' Activity Start date is not required.
        '691	REM - AS/SVE System
        '692	REM - Dual Phase System
        '695	REM - Pump & Treat System
        If ugActivitiesandDocuments.ActiveRow.Band.Index = 0 Then
            If Not (ugActivitiesandDocuments.ActiveRow.Cells("Activity_Type_ID").Value = 691 Or _
                ugActivitiesandDocuments.ActiveRow.Cells("Activity_Type_ID").Value = 692 Or _
                 ugActivitiesandDocuments.ActiveRow.Cells("Activity_Type_ID").Value = 693) Then
                If IsDBNull(ugActivitiesandDocuments.ActiveRow.Cells("Start_Date").Value) Then
                    MessageBox.Show("Activity Start Date is required for adding a Document.")
                    Exit Sub
                End If
            End If
        End If

        Dim frmAddDocument As New Document(oLustEvent)
        Dim MyFrm As MusterContainer

        Try
            MyFrm = MdiParent
            Me.Tag = 0

            If ugActivitiesandDocuments.ActiveRow.Band.Index = 0 Then
                frmAddDocument.CallingForm = Me
                frmAddDocument.Mode = 0 ' Add
                frmAddDocument.EventActivityID = ugActivitiesandDocuments.ActiveRow.Cells("Event_Activity_ID").Value
                frmAddDocument.EventOwnerID = oOwner.ID
                frmAddDocument.EventDocumentID = 0
                frmAddDocument.TFStatus = oLustEvent.MGPTFStatus
                frmAddDocument.ActivityName = ugActivitiesandDocuments.ActiveRow.Cells("Activity").Value
                If Not IsDBNull(ugActivitiesandDocuments.ActiveRow.Cells("Start_Date").Value) Then
                    frmAddDocument.StartDate = ugActivitiesandDocuments.ActiveRow.Cells("Start_Date").Value
                End If

                frmAddDocument.ShowDialog()
                UpdateEventDates()
                GetLustActivitiesForEvent(oLustEvent.ID)
                GetLustRemediationForEvent(oLustEvent.ID)
                If Me.Tag = 1 Then
                    ShowDocumentList()
                End If
                If Me.Tag = 2 Then
                    oLustEvent.EventStatus = 625
                    oLustEvent.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    oLustEvent.CloseOpenNFAActivities(oLustEvent.ID)
                    ShowDocumentList()
                End If


                MyFrm.RefreshCalendarInfo()
                MyFrm.LoadDueToMeCalendar()
                MyFrm.LoadToDoCalendar()
                Me.SetupModifyLustEventForm(True, False)
            Else
                MsgBox("You must select an Activity to associate the document with.")
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Add Document" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmAddDocument = Nothing
        End Try
    End Sub
    Private Sub btnModifyDocument_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifyDocument.Click
        ModifyDocument()
    End Sub
    Private Sub ugActivitiesandDocuments_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugActivitiesandDocuments.DoubleClick
        If Me.Cursor Is Cursors.WaitCursor Then
            Exit Sub
        End If
        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        If bolAllowDocPop OrElse (ugActivitiesandDocuments.ActiveRow.Band.Index = 1 AndAlso _
        (String.Format("{0} ", ugActivitiesandDocuments.ActiveRow.Cells("Document").Value.ToString.ToUpper).IndexOf(" SOW ") > -1 Or _
         String.Format("{0} ", ugActivitiesandDocuments.ActiveRow.Cells("Document").Value.ToString.ToUpper).IndexOf(" SOW/CE ") > -1) _
         ) Then
            If ugActivitiesandDocuments.ActiveRow.Band.Index = 0 Then
                ModifyActivity()
            Else
                ModifyDocument()
            End If
        Else
            MsgBox("Update not allowed for user.")
        End If
    End Sub
    Private Sub ugActivitiesandDocuments_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugActivitiesandDocuments.AfterRowActivate
        If (bolPMHead Or MusterContainer.AppUser.HEAD_ADMIN) Or cmbProjectManager.Text.ToUpper = MusterContainer.AppUser.Name.ToUpper Then
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            ugRow = ugActivitiesandDocuments.ActiveRow

            btnModifyActivity.Enabled = False
            btnAddDocument.Enabled = False
            btnModifyDocument.Enabled = False
            btnDeleteActivity.Enabled = False
            btnDeleteDocument.Enabled = False

            If Not ugRow Is Nothing Then
                btnModifyActivity.Enabled = True
                btnDeleteActivity.Enabled = True
                If ugRow.Band.Index = 1 Then
                    ugRow = ugRow.ParentRow
                End If
                If ugRow.Cells("ActivityDocsRelCount").Value > 0 Then
                    btnAddDocument.Enabled = True
                    btnModifyDocument.Enabled = True
                    btnDeleteDocument.Enabled = True
                End If
            End If
        End If
    End Sub
    Private Sub ugActivitiesandDocuments_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugActivitiesandDocuments.BeforeCellActivate
        Try
            lblActivitiesDocumentsCellDesc.Text = ""
            lblActivitiesDocumentsCellDesc.BackColor = SystemColors.Control
            If e.Cell.Band.Index = 0 Then
                If e.Cell.Row.CellAppearance.BackColor.ToString = Color.Red.ToString Then
                    lblActivitiesDocumentsCellDesc.Text = "Days Since Start > Act By"
                    lblActivitiesDocumentsCellDesc.BackColor = e.Cell.Row.CellAppearance.BackColor
                ElseIf e.Cell.Row.CellAppearance.BackColor.ToString = Color.Yellow.ToString Then
                    lblActivitiesDocumentsCellDesc.Text = "Days Since Start > Warn"
                    lblActivitiesDocumentsCellDesc.BackColor = e.Cell.Row.CellAppearance.BackColor
                End If
            Else
                If e.Cell.Column.Key.ToUpper = "DUE" And e.Cell.Appearance.BackColor.ToString = Color.Yellow.ToString And e.Cell.Row.Cells("RevisionDate").Value Is DBNull.Value Then
                    lblActivitiesDocumentsCellDesc.Text = "No RevisionDate Date"
                    lblActivitiesDocumentsCellDesc.BackColor = e.Cell.Appearance.BackColor
                ElseIf e.Cell.Row.CellAppearance.BackColor.ToString = Color.Orange.ToString Then
                    If e.Cell.Row.Cells("Extension").Value Is DBNull.Value Then
                        lblActivitiesDocumentsCellDesc.Text = "Due < Today AND No Closed Date, Received Date, To Financial Date AND No Extension Date"
                        lblActivitiesDocumentsCellDesc.BackColor = e.Cell.Row.CellAppearance.BackColor
                    ElseIf e.Cell.Row.Cells("Extension").Value < Now.Date Then
                        lblActivitiesDocumentsCellDesc.Text = "Due < Today AND No Closed Date, Received Date, To Financial Date AND Extension < Today"
                        lblActivitiesDocumentsCellDesc.BackColor = e.Cell.Row.CellAppearance.BackColor
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Displaying Activities/Documents Cell Description : " & vbCrLf & ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            lblActivitiesDocumentsCellDesc.Text = ""
            lblActivitiesDocumentsCellDesc.BackColor = SystemColors.Control
        End Try
    End Sub

    Private Sub ugRemediationSystem_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugRemediationSystem.DoubleClick
        If bolTabClick = False Then
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        End If

        ModifyRemediation()
    End Sub

    Private Sub btnModifyRemediationSystem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifyRemediationSystem.Click
        ModifyRemediation()
    End Sub
#End Region

#Region "Lust Event Form Operations"

    Private Sub ShowDocumentList()

        Dim frmLtr As Letters
        frmLtr = New Letters(1)
        frmLtr.MdiParent = Me.MdiParent
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

    Friend Sub SetupAddLustEventForm()
        Dim tmpdate As Date
        Try
            nCurrentEventID = -1
            ugActivitiesandDocuments.Width = pnlActivitiesDocumentsDetails.Width - 50
            tbPageLUSTEvent.Enabled = True
            If Not tbCntrlTechnical.TabPages.Contains(tbPageLUSTEvent) Then
                tbCntrlTechnical.TabPages.Add(tbPageLUSTEvent)
            End If

            If oLustEvent.IsDirty Then
                If MsgBox("Do you wish to save changes from previously viewed event?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    ProcessSaveEvent()
                End If
            End If
            oLustEvent.FacilityID = oOwner.Facilities.ID

            ClearLUSTForm()
            bolLoading = True

            ResetChecklist()
            tbCntrlTechnical.SelectedTab = tbPageLUSTEvent
            PopulateLustEventLookups()

            cmbReleaseStatus.SelectedText = "Suspected"
            cmbEventStatus.SelectedText = "Open"
            cmbMGPTFStatus.SelectedText = "EUD"

            UIUtilsGen.SetDatePickerValue(dtStartDate, Now.Date)
            UIUtilsGen.SetDatePickerValue(dtDateofReport, "01/01/0001")
            UIUtilsGen.SetDatePickerValue(dtCompAssDate, "01/01/0001")
            oLustEvent.ReportDate = tmpdate
            oLustEvent.CompAssDate = tmpdate
            oLustEvent.Started = dtStartDate.Value.Date
            oLustEvent.EventStarted = dtStartDate.Value.Date
            oLustEvent.ReleaseStatus = cmbReleaseStatus.SelectedValue
            oLustEvent.EventStatus = cmbEventStatus.SelectedValue
            oLustEvent.MGPTFStatus = cmbMGPTFStatus.SelectedValue
            oLustEvent.PM = PMHead_StaffID

            ' ===============
            chkSoilBTEX.Enabled = False
            chkSoilPAH.Enabled = False
            chkSoilTPH.Enabled = False

            chkGroundWaterBTEX.Enabled = False
            chkGroundWaterPAH.Enabled = False
            chkGroundWaterTPH.Enabled = False

            chkFreeProductDiesel.Enabled = False
            chkFreeProductGasoline.Enabled = False
            chkFreeProductKerosene.Enabled = False
            chkFreeProductUnKnown.Enabled = False
            chkFreeProductWasteOil.Enabled = False

            chkVaporBTEX.Enabled = False
            chkVaporPAH.Enabled = False

            ' ===============

            pnlFundsEligibility.Visible = False
            pnlFundsEligibilityDetails.Visible = False
            pnlActivitiesDocuments.Visible = False
            pnlActivitiesDocumentsDetails.Visible = False
            pnlRemediationSystems.Visible = False
            pnlRemediationSystemsDetails.Visible = False
            pnlERACandContacts.Visible = False
            pnlERACandContactsDetails.Visible = False
            'pnlPermits.Visible = False
            'pnlPermitsDetails.Visible = False


            pnlReleaseInfo.Visible = True
            lblReleaseInfoDisplay.Text = "-"
            PnlReleaseInfoDetails.Visible = True

            lblEventInfoDisplay.Text = "-"
            PnlEventInfoDetails.Visible = True

            lblCommentsDisplay.Text = "-"
            pnlComments.Visible = True
            pnlCommentsDetails.Visible = True

            pnlERACandContactsDetails.Visible = False
            lblERACandContactsDisplay.Text = "+"

            GetLustTanksAndPipesForFacility(0)

            dtStartDate.Enabled = True
            dtDateofReport.Enabled = True
            dtCompAssDate.Enabled = True
            If PMHead_StaffID = 0 Then
                MsgBox("No PM-Head Identified to the system")
                cmbProjectManager.SelectedIndex = 1
            End If
            SetPageSecurity(True)
            LoadContacts(ugERACContacts, 0, 7)
            btnGoToFinancial.Enabled = False
            'If oLustEvent.ID = 0 Then
            '    Me.btnDeleteLUSTEvent.Enabled = False
            'End If
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Setting Up Add Lust Event Form" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ProcessSystemFlags()
        Dim mc As MusterContainer
        Try
            mc = CType(Me.MdiParent, MusterContainer)
            If nLastEventID <> 0 Then
                mc.FlagsChanged(oLustEvent.FacilityID, UIUtilsGen.EntityTypes.Facility, "Technical", Me.Text, nLastEventID, UIUtilsGen.EntityTypes.LUST_Event)
                ''ElseIf oLustEvent.ID <> 0 Then
                ''    mc.FlagsChanged(oLustEvent.FacilityID, UIUtilsGen.EntityTypes.Facility, "Technical", Me.Text, oLustEvent.ID, UIUtilsGen.EntityTypes.LUST_Event)
            Else
                mc.FlagsChanged(oLustEvent.FacilityID, UIUtilsGen.EntityTypes.Facility, "Technical", Me.Text)
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing System Flags" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub SetupAfterAdd()
        Try
            pnlReleaseInfo.Visible = True
            pnlFundsEligibility.Visible = True
            pnlActivitiesDocuments.Visible = True
            pnlComments.Visible = True
            pnlRemediationSystems.Visible = True
            pnlERACandContacts.Visible = True
            'pnlPermits.Visible = True
            pnlFundsEligibilityDetails.Visible = False
            lblReleaseInfoDisplay.Text = "-"
            lblActivitiesDocumentsDisplay.Text = "+"
            lblRemediationSystemsDisplay.Text = "+"
            lblERACandContactsDisplay.Text = "+"
            'lblPermitsDisplay.Text = "+"
            lblFundsEligibilityDisplay.Text = "+"
            lblEventID.Text = "Event #:"
            lblEventIDValue.Text = oLustEvent.EVENTSEQUENCE
            lblEventIDValue.Visible = True
            lblEventCountValue.Visible = False
            GetLustTanksAndPipesForFacility(oLustEvent.ID)
            GetLustActivitiesForEvent(oLustEvent.ID)
            ProcessSystemFlags()
            If cmbProjectManager.Text.ToUpper = MusterContainer.AppUser.Name.ToUpper Then
                'pnlActivitiesDocumentsDetails.Enabled = True
                btnAddActivity.Enabled = True
                btnModifyActivity.Enabled = True
                btnDeleteActivity.Enabled = True
                btnAddDocument.Enabled = True
                btnDeleteDocument.Enabled = True
                btnModifyDocument.Enabled = True
            Else
                'pnlActivitiesDocumentsDetails.Enabled = False
                btnAddActivity.Enabled = False
                btnModifyActivity.Enabled = False
                btnDeleteActivity.Enabled = False
                btnAddDocument.Enabled = False
                btnDeleteDocument.Enabled = False
                btnModifyDocument.Enabled = False
            End If



            If cmbReleaseStatus.Text = "Confirmed" Then
                'disable FundsEligibility panel
                pnlFundsEligibility.Visible = True
                pnlFundsEligibilityQuestions.Visible = True
                'pnlFundsEligibilityDetails.Visible = True
                'lblFundsEligibilityDisplay.Text = "-"
                If (bolPMHead Or MusterContainer.AppUser.HEAD_ADMIN) Then
                    pnlTFAssess.Enabled = True
                Else
                    pnlTFAssess.Enabled = False
                End If
            Else
                '2 lines to disable Funds Eligibility
                pnlFundsEligibility.Visible = False
                pnlFundsEligibilityQuestions.Visible = False
                pnlFundsEligibilityDetails.Visible = False
                lblFundsEligibilityDisplay.Text = "+"
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Setting Up After Add" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Friend Sub SetupModifyLustEventForm(Optional ByVal bolFromSave As Boolean = False, Optional ByVal bolInitialCreation As Boolean = False)
        Dim EventToLoad As Int64
        Try
            ugActivitiesandDocuments.Width = pnlActivitiesDocumentsDetails.Width - 50

            ugRemediationSystem.Width = pnlRemediationSystemsDetails.Width - 50
            tbPageLUSTEvent.Enabled = True
            If Not tbCntrlTechnical.TabPages.Contains(tbPageLUSTEvent) Then
                tbCntrlTechnical.TabPages.Add(tbPageLUSTEvent)
            End If

            If Not bolFromSave Then
                EventToLoad = dgLUSTEvents.ActiveRow.Cells("EVENT_ID").Text
                If oLustEvent.IsDirty Then
                    If MsgBox("Do you wish to save changes from previously viewed event?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        ProcessSaveEvent()
                    End If
                End If
                oLustEvent.Retrieve(EventToLoad)
                nLastEventID = EventToLoad
            Else
                nLastEventID = oLustEvent.ID
            End If
            nCurrentEventID = nLastEventID

            bolLoading = True
            If Not bolFromSave Then
                ResetChecklist()
            End If

            tbCntrlTechnical.SelectedTab = tbPageLUSTEvent

            lblEventID.Text = "Event #:"
            lblEventIDValue.Text = oLustEvent.EVENTSEQUENCE
            lblEventIDValue.Visible = True
            lblEventCountValue.Visible = False
            PopulateLustEventLookups()

            pnlReleaseInfo.Visible = True
            pnlActivitiesDocuments.Visible = True
            pnlComments.Visible = True
            pnlRemediationSystems.Visible = True
            pnlERACandContacts.Visible = True
            'pnlPermits.Visible = True

            pnlActivitiesDocumentsDetails.Visible = True
            pnlCommentsDetails.Visible = True
            pnlRemediationSystemsDetails.Visible = True
            pnlERACandContactsDetails.Visible = True
            'pnlPermitsDetails.Visible = True

            pnlERACandContactsDetails.Visible = True
            lblERACandContactsDisplay.Text = "-"

            ' ===============
            chkSoilBTEX.Enabled = False
            chkSoilPAH.Enabled = False
            chkSoilTPH.Enabled = False

            chkGroundWaterBTEX.Enabled = False
            chkGroundWaterPAH.Enabled = False
            chkGroundWaterTPH.Enabled = False

            chkFreeProductDiesel.Enabled = False
            chkFreeProductGasoline.Enabled = False
            chkFreeProductKerosene.Enabled = False
            chkFreeProductUnKnown.Enabled = False
            chkFreeProductWasteOil.Enabled = False

            chkVaporBTEX.Enabled = False
            chkVaporPAH.Enabled = False

            LoadLustEvent()

            GetLustTanksAndPipesForFacility(oLustEvent.ID)

            LoadTFChecklist()
            '    If Not bolFromSave Then
            '    Else
                'LoadTFChecklist()
            '   End If


            GetLustActivitiesForEvent(oLustEvent.ID)
            GetLustRemediationForEvent(oLustEvent.ID)
            If oLustEvent.TankandPipe <> String.Empty Then
                EventTankToggle()
            End If

            pnlFundsEligibilityDetails.Visible = False
            lblFundsEligibilityDisplay.Text = "+"

            If cmbReleaseStatus.Text = "Confirmed" Then
                cmbSuspectedSource.Enabled = False
                lblReleaseInfoDisplay.Text = "+"
                PnlReleaseInfoDetails.Visible = False

                'disable FundsEligibility panel
                pnlFundsEligibility.Visible = True
                pnlFundsEligibilityQuestions.Visible = True
                pnlFundsEligibilityDetails.Enabled = True
                'pnlFundsEligibilityDetails.Visible = True
                'lblFundsEligibilityDisplay.Text = "-"

                If bolPMHead Or MusterContainer.AppUser.HEAD_ADMIN Then
                    pnlTFAssess.Enabled = True
                Else
                    pnlTFAssess.Enabled = False
                End If
            Else
                cmbSuspectedSource.Enabled = True
                lblReleaseInfoDisplay.Text = "-"
                PnlReleaseInfoDetails.Visible = True

                '3 lines to disable Funds Eligibility
                pnlFundsEligibility.Visible = False
                pnlFundsEligibilityDetails.Visible = False
                pnlFundsEligibilityQuestions.Visible = False
                lblFundsEligibilityDisplay.Text = "+"

                pnlFundsEligibilityDetails.Enabled = False
            End If

            If bolInitialCreation Then
                lblReleaseInfoDisplay.Text = "+"
                PnlReleaseInfoDetails.Visible = False
            End If

            ' ===============
            lblCommentsDisplay.Text = "-"
            pnlCommentsDetails.Visible = True

            If bolInitialCreation Then
                lblActivitiesDocumentsDisplay.Text = "+"
                pnlActivitiesDocumentsDetails.Visible = False
            Else
                lblActivitiesDocumentsDisplay.Text = "-"
                pnlActivitiesDocumentsDetails.Visible = True
            End If

            If bolInitialCreation Then
                lblEventInfoDisplay.Text = "-"
                PnlEventInfoDetails.Visible = True
            Else
                lblEventInfoDisplay.Text = "+"
                PnlEventInfoDetails.Visible = False
            End If

            lblRemediationSystemsDisplay.Text = "+"
            pnlRemediationSystemsDetails.Visible = False

            lblERACandContactsDisplay.Text = "+"
            pnlERACandContactsDetails.Visible = False

            'lblPermitsDisplay.Text = "+"
            'pnlPermitsDetails.Visible = False


            If Not (bolPMHead Or MusterContainer.AppUser.HEAD_ADMIN) Then
                dtStartDate.Enabled = False
                dtDateofReport.Enabled = False
                dtCompAssDate.Enabled = True

            Else
                dtStartDate.Enabled = True
                dtDateofReport.Enabled = True
                dtCompAssDate.Enabled = True
            End If

            SetPMHistory()

            oFinancialEvent.Retrieve_ByTechnicalEventID(oLustEvent.ID)

            If oFinancialEvent.ID > 0 Then
                btnGoToFinancial.Enabled = True
            Else
                btnGoToFinancial.Enabled = False
            End If

            ' ===============
            bolLoading = False
            If IsNothing(Me.MdiParent) = False Then
                ProcessSystemFlags()
            End If

            'If cmbProjectManager.Text.ToUpper = MusterContainer.AppUser.Name.ToUpper Or bolPMHead Then
            '    pnlActivitiesDocumentsDetails.Enabled = True
            'Else
            '    pnlActivitiesDocumentsDetails.Enabled = False
            'End If
            SetPageSecurity(False)
            btnViewModifyComment.Focus()
            'If oLustEvent.ID <> 0 Then
            '    Me.btnDeleteLUSTEvent.Enabled = True
            'End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Setting up Modify Lust Event Form" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Friend Sub ClearLUSTForm()
        Try
            bolLoading = True
            chkFacilityLeakDetection.Checked = False
            chkFailedPTT.Checked = False
            chkFreeProduct.Checked = False
            chkFreeProductDiesel.Checked = False
            chkFreeProductGasoline.Checked = False
            chkFreeProductKerosene.Checked = False
            chkFreeProductUnKnown.Checked = False
            chkFreeProductWasteOil.Checked = False
            chkGroundWater.Checked = False
            chkGroundWaterBTEX.Checked = False
            chkGroundWaterPAH.Checked = False
            chkGroundWaterTPH.Checked = False
            chkGWWell.Checked = False
            chkInspection.Checked = False
            chkInventoryShortage.Checked = False
            chkSoil.Checked = False
            chkSoilBTEX.Checked = False
            chkSoilPAH.Checked = False
            chkSoilTPH.Checked = False
            chkSurfaceSheen.Checked = False
            chkTankClosure.Checked = False
            chkToCFreeProduct.Checked = False
            chkToCVapor.Checked = False
            chkVaporBTEX.Checked = False
            chkVaporPAH.Checked = False
            chkVapors.Checked = False

            cmbSuspectedSource.Enabled = True
            txtRelatedSites.Text = String.Empty
            UIUtilsGen.ClearFields(tbPageLUSTEvent)
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Clear Form" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
        End Try
    End Sub
    Private Sub LoadLustEvent()
        Dim xUser As New MUSTER.BusinessLogic.pUser
        Dim xFind As Integer
        Dim MyFrm As MusterContainer
        Try
            If nSavedEventID <> oLustEvent.ID Then
                oLustEvent.MarkToDoCompleted_ByDesc(oLustEvent.ID, "Transferred Lust Event for ID")
                If IsNothing(MdiParent) = False Then
                    MyFrm = MdiParent
                    MyFrm.RefreshCalendarInfo()
                    MyFrm.LoadDueToMeCalendar()
                    MyFrm.LoadToDoCalendar()
                End If
            End If
            Me.Text = "Technical LUST Events - Facility #" & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
            nOriginalPM = oLustEvent.PM
            lblProjectManager.Visible = False
            cmbProjectManager.Visible = True
            If oLustEvent.PM > 0 Then
                xUser.Retrieve(oLustEvent.PM)
                xFind = cmbProjectManager.FindString(xUser.Name)
                If xFind > -1 Then
                    cmbProjectManager.SelectedValue = oLustEvent.PM
                Else
                    If UCase(oLustEvent.TechnicalStatusDesc) <> UCase("Closed") Then
                        If MsgBox("PM Assigned To This Event No Longer Valid, Do You Want To Reassign the PM Now?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                            lblProjectManager.Text = xUser.Name
                            lblProjectManager.Visible = True
                            cmbProjectManager.Visible = False
                        End If
                    Else
                        lblProjectManager.Text = xUser.Name
                        lblProjectManager.Visible = True
                        cmbProjectManager.Visible = False
                    End If
                End If
            End If

            If bolPMHead Or cmbProjectManager.Text.ToUpper = MusterContainer.AppUser.Name.ToUpper Or (MusterContainer.AppUser.HEAD_ADMIN) Then
                btnDeleteActivity.Visible = True
                btnDeleteDocument.Visible = True
            Else
                btnDeleteActivity.Visible = False
                btnDeleteDocument.Visible = False
            End If
            cmbEventStatus.SelectedValue = oLustEvent.EventStatus
            cmbMGPTFStatus.SelectedValue = oLustEvent.MGPTFStatus
            If oLustEvent.Priority > 0 Then
                cmbPriority.SelectedValue = oLustEvent.Priority
            Else
                cmbPriority.SelectedIndex = -1
                cmbPriority.SelectedIndex = -1
            End If
            UIUtilsGen.SetDatePickerValue(dtStartDate, oLustEvent.Started)
            UIUtilsGen.SetDatePickerValue(dtDateofReport, oLustEvent.ReportDate)
            UIUtilsGen.SetDatePickerValue(dtCompAssDate, oLustEvent.CompAssDate)

            cmbReleaseStatus.SelectedValue = oLustEvent.ReleaseStatus
            If oLustEvent.SuspectedSource > 0 Then
                SetDropDown(cmbSuspectedSource, oLustEvent.SuspectedSource)
            Else
                cmbSuspectedSource.SelectedIndex = 0
            End If
            If cmbReleaseStatus.SelectedText.ToUpper = "CONFIRMED" Then
                cmbSuspectedSource.Enabled = False
            End If
            txtRelatedSites.Text = oLustEvent.RelatedSites
            UIUtilsGen.SetDatePickerValue(dtLastLDR, oLustEvent.LastLDR)
            UIUtilsGen.SetDatePickerValue(dtLastPTT, oLustEvent.LastPTT)
            UIUtilsGen.SetDatePickerValue(dtLastGWS, oLustEvent.LastGWS)
            UIUtilsGen.SetDatePickerValue(dtConfirmedOn, oLustEvent.Confirmed)

            UIUtilsGen.SetDatePickerValue(dtUSTChiefOn, oLustEvent.UST_CHIEF_DATE)
            UIUtilsGen.SetDatePickerValue(dtPMHeadOn, oLustEvent.PM_HEAD_DATE)
            UIUtilsGen.SetDatePickerValue(dtCommissionOn, oLustEvent.COMMISSION_DATE)
            UIUtilsGen.SetDatePickerValue(dtOPCHeadOn, oLustEvent.OPC_HEAD_DATE)

            txtCommissionBy.Text = oLustEvent.COMMISSION_BY
            txtUSTChiefBy.Text = oLustEvent.UST_CHIEF_BY
            txtPMHeadBy.Text = oLustEvent.PM_HEAD_BY
            txtOPCHeadBy.Text = oLustEvent.OPC_HEAD_BY

            cmbIdentifiedBy.SelectedValue = oLustEvent.IDENTIFIEDBY
            cmbLocation.SelectedValue = oLustEvent.Location
            cmbExtent.SelectedValue = oLustEvent.Extent
            cmbCause.SelectedValue = oLustEvent.Cause

            chkSoil.Checked = oLustEvent.TOCSOIL
            If chkSoil.Checked Then
                chkSoilBTEX.Checked = oLustEvent.SOILBTEX
                chkSoilPAH.Checked = oLustEvent.SOILPAH
                chkSoilTPH.Checked = oLustEvent.SOILTPH

                chkSoilBTEX.Enabled = True
                chkSoilPAH.Enabled = True
                chkSoilTPH.Enabled = True
            End If


            chkGroundWater.Checked = oLustEvent.TOCGROUNDWATER
            If chkGroundWater.Checked Then
                chkGroundWaterBTEX.Checked = oLustEvent.GWBTEX
                chkGroundWaterPAH.Checked = oLustEvent.GWPAH
                chkGroundWaterTPH.Checked = oLustEvent.GWTPH

                chkGroundWaterBTEX.Enabled = True
                chkGroundWaterPAH.Enabled = True
                chkGroundWaterTPH.Enabled = True
            End If

            chkToCFreeProduct.Checked = oLustEvent.FREEPRODUCT
            If chkToCFreeProduct.Checked Then
                chkFreeProductDiesel.Checked = oLustEvent.FPDIESEL
                chkFreeProductGasoline.Checked = oLustEvent.FPGASOLINE
                chkFreeProductKerosene.Checked = oLustEvent.FPKEROSENE
                chkFreeProductUnKnown.Checked = oLustEvent.FPUNKNOWN
                chkFreeProductWasteOil.Checked = oLustEvent.FPWASTEOIL

                chkFreeProductDiesel.Enabled = True
                chkFreeProductGasoline.Enabled = True
                chkFreeProductKerosene.Enabled = True
                chkFreeProductUnKnown.Enabled = True
                chkFreeProductWasteOil.Enabled = True
            End If

            chkToCVapor.Checked = oLustEvent.TOCVAPOR
            If chkToCVapor.Checked Then
                chkVaporBTEX.Checked = oLustEvent.VAPORBTEX
                chkVaporPAH.Checked = oLustEvent.VAPORPAH

                chkVaporBTEX.Enabled = True
                chkVaporPAH.Enabled = True
            End If

            'If chkFacilityLeakDetection.Tag = oLustEvent.HowDiscoveredID Then
            '    chkFacilityLeakDetection.Checked = True
            'End If

            'If chkSurfaceSheen.Tag = oLustEvent.HowDiscoveredID Then
            '    chkSurfaceSheen.Checked = True
            'End If

            'If chkGWWell.Tag = oLustEvent.HowDiscoveredID Then
            '    chkGWWell.Checked = True
            'End If

            'If chkGWContamination.Tag = oLustEvent.HowDiscoveredID Then
            '    chkGWContamination.Checked = True
            'End If

            'If chkVapors.Tag = oLustEvent.HowDiscoveredID Then
            '    chkVapors.Checked = True
            'End If

            'If chkFreeProduct.Tag = oLustEvent.HowDiscoveredID Then
            '    chkFreeProduct.Checked = True
            'End If

            'If chkSoilContamination.Tag = oLustEvent.HowDiscoveredID Then
            '    chkSoilContamination.Checked = True
            'End If

            'If chkFailedPTT.Tag = oLustEvent.HowDiscoveredID Then
            '    chkFailedPTT.Checked = True
            'End If

            'If chkInventoryShortage.Tag = oLustEvent.HowDiscoveredID Then
            '    chkInventoryShortage.Checked = True
            'End If

            'If chkTankClosure.Tag = oLustEvent.HowDiscoveredID Then
            '    chkTankClosure.Checked = True
            'End If

            'If chkInspection.Tag = oLustEvent.HowDiscoveredID Then
            '    chkInspection.Checked = True
            'End If

            chkFacilityLeakDetection.Checked = oLustEvent.HowDiscFacLD
            chkSurfaceSheen.Checked = oLustEvent.HowDiscSurfaceSheen
            chkGWWell.Checked = oLustEvent.HowDiscGWWell
            chkGWContamination.Checked = oLustEvent.HowDiscGWContamination
            chkVapors.Checked = oLustEvent.HowDiscVapors
            chkFreeProduct.Checked = oLustEvent.HowDiscFreeProduct
            chkSoilContamination.Checked = oLustEvent.HowDiscSoilContamination
            chkFailedPTT.Checked = oLustEvent.HowDiscFailedPTT
            chkInventoryShortage.Checked = oLustEvent.HowDiscInventoryShortage
            chkTankClosure.Checked = oLustEvent.HowDiscTankClosure
            chkInspection.Checked = oLustEvent.HowDiscInspection

            Select Case oLustEvent.COMMISSION_ASSESS
                Case 1
                    chkCommissionYes.Checked = True
                Case 2
                    chkCommissionNo.Checked = True
            End Select

            Select Case oLustEvent.OPC_HEAD_ASSESS
                Case 1
                    chkOPCHeadYes.Checked = True
                Case 2
                    chkOPCHeadNo.Checked = True
            End Select

            Select Case oLustEvent.PM_HEAD_ASSESS
                Case 1
                    chkPMHeadYes.Checked = True
                Case 3
                    chkPMHeadUndecided.Checked = True
            End Select

            Select Case oLustEvent.UST_CHIEF_ASSESS
                Case 1
                    Me.chkUSTChiefYes.Checked = True
                Case 2
                    Me.chkUSTChiefNo.Checked = True
                Case 3
                    Me.chkUSTChiefUndecided.Checked = True
            End Select
            If chkPMHeadYes.Checked = False And chkUSTChiefYes.Checked = False And chkOPCHeadYes.Checked = False And chkCommissionYes.Checked = False Then
                pnlFundsEligibilityQuestions.Enabled = True
            Else
                pnlFundsEligibilityQuestions.Enabled = False
            End If
            chkforHeadofOPC.Checked = oLustEvent.FOR_OPC_HEAD
            chkforCommission.Checked = oLustEvent.FOR_COMMISSION
            txtEligibilityComments.Text = oLustEvent.ELIGIBITY_COMMENTS

            If oLustEvent.ModifiedBy Is Nothing Or oLustEvent.ModifiedBy = String.Empty Then
                lblOwnerLastEditedBy.Text = "Last Edited By : " & oLustEvent.CreatedBy.ToString()
            Else
                lblOwnerLastEditedBy.Text = "Last Edited By : " & oLustEvent.ModifiedBy.ToString()
            End If
            lblOwnerLastEditedOn.Text = "Last Edited On : " & IIf(oLustEvent.ModifiedOn = CDate("01/01/0001"), oLustEvent.CreatedOn.ToString, oLustEvent.ModifiedOn.ToString)

            pCompany.Retrieve(oLustEvent.IRAC)
            txtIRAC.Text = pCompany.COMPANY_NAME

            pCompany.Retrieve(oLustEvent.ERAC)
            txtERAC.Text = pCompany.COMPANY_NAME

            LoadContacts(ugERACContacts, oLustEvent.ID, 7)

            If oLustEvent.ID <= 0 And oLustEvent.ID >= -100 Then
                CommentsMaintenance(, , True, True)
            Else
                CommentsMaintenance(, , True)
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Lust Event " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally

        End Try


    End Sub

    Private Sub SetDropDown(ByRef ctlDropDown As ComboBox, ByVal Value As String)
        Try

            ctlDropDown.SelectedValue = Value
        Catch ex As Exception
            'Nothing
        End Try
    End Sub
    Private Sub EventTankToggle()
        Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim bolActiveRow As Boolean

        If btnEvtTankToggle.Text = "Show Event Only" Then
            For Each ugrow In ugTankandPipes.Rows
                bolActiveRow = False
                If ugrow.Band.Index = 0 Then
                    ChildBand = ugrow.ChildBands(0)
                    For Each Childrow In ChildBand.Rows
                        If Childrow.Cells("Included").Value = 0 Then
                            Childrow.Hidden = True
                        Else
                            bolActiveRow = True
                        End If
                    Next
                    If ugrow.Cells("Included").Value = 0 And bolActiveRow = False Then
                        ugrow.Hidden = True
                    End If
                End If
            Next
            btnEvtTankToggle.Text = "Show All"
        Else
            For Each ugrow In ugTankandPipes.Rows
                ChildBand = ugrow.ChildBands(0)
                For Each Childrow In ChildBand.Rows
                    Childrow.Hidden = False
                Next
                ugrow.Hidden = False
            Next
            btnEvtTankToggle.Text = "Show Event Only"
        End If
    End Sub

    Private Sub GenerateFlag_OnAdd()


    End Sub

    Private Sub GenerateActivity_NewSite()
        Dim oLustEventActivity As New MUSTER.BusinessLogic.pLustEventActivity
        Dim MyFrm As MusterContainer

        Try
            MyFrm = MdiParent

            oLustEventActivity.Add(New MUSTER.Info.LustActivityInfo(0, _
                oLustEvent.ID, _
                Now.Date, _
                CDate("01/01/0001"), _
                CDate("01/01/0001"), _
                CDate("01/01/0001"), _
                CDate("01/01/0001"), _
                682, _
                MusterContainer.AppUser.ID, _
                Now.Date, _
                String.Empty, _
                CDate("01/01/0001"), _
                0, 0))

            oLustEventActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            oLustEvent.CalendarEntries(cmbProjectManager.Text, "New Lust Event for ID: " & oLustEvent.FacilityID & " Event " & oLustEvent.EVENTSEQUENCE, True, False, String.Empty, UserID, Now.Date, Now.Date, 0, oLustEventActivity.ActivityID)
            MyFrm.RefreshCalendarInfo()
            MyFrm.LoadDueToMeCalendar()
            MyFrm.LoadToDoCalendar()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing:  Generate New Site Activity" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub

    Private Sub TestForConfirmedSpill()
        Dim dtTempDate As Date
        Dim bolConfirmed As Boolean = False
        Dim nReleaseStatus As Integer

        If bolLoading Then Exit Sub

        If dtConfirmedOn.Checked = True And Date.Compare(dtConfirmedOn.Value.Date, dtTempDate) <> 0 Then
            bolConfirmed = True
        End If

        If cmbLocation.SelectedValue <> 0 Then
            bolConfirmed = True
        End If

        If cmbExtent.SelectedValue <> 0 Then
            bolConfirmed = True
        End If

        If oLustEvent.TankandPipe <> String.Empty Then
            bolConfirmed = True
        End If

        If chkToCVapor.Checked Or chkToCFreeProduct.Checked Or chkSoil.Checked Or chkGroundWater.Checked Then
            bolConfirmed = True
        End If

        If chkSurfaceSheen.Checked Or chkGWContamination.Checked Or chkGWWell.Checked Or chkSoilContamination.Checked Or chkFailedPTT.Checked Or chkFreeProduct.Checked Or chkTankClosure.Checked Then
            bolConfirmed = True
        End If

        If bolConfirmed Then
            nReleaseStatus = oLustEvent.ReleaseStatus
            cmbReleaseStatus.SelectedValue = 623
            If nReleaseStatus <> cmbReleaseStatus.SelectedValue Then
                bolNewRelease = True
            End If
        Else
            bolNewRelease = False
        End If
    End Sub


    Private Sub ResetChecklist()

        Me.chkQuestion1Yes.Checked = False
        Me.chkQuestion1No.Checked = False

        Me.chkQuestion2Yes.Checked = False
        Me.chkQuestion2No.Checked = False

        Me.chkQuestion5Yes.Checked = False
        Me.chkQuestion5No.Checked = False
        Me.chkQuestion5NA.Checked = False

        Me.chkQuestion6Yes.Checked = False
        Me.chkQuestion6No.Checked = False
        Me.chkQuestion6NA.Checked = False

        Me.chkQuestion8Yes.Checked = False
        Me.chkQuestion8No.Checked = False
        Me.chkQuestion8NA.Checked = False

        Me.chkQuestion10Yes.Checked = False
        Me.chkQuestion10No.Checked = False
        Me.chkQuestion10NA.Checked = False

        Me.chkQuestion14Yes.Checked = False
        Me.chkQuestion14No.Checked = False
        Me.chkQuestion14NA.Checked = False

        Me.chkQuestion15Yes.Checked = False
        Me.chkQuestion15No.Checked = False
        Me.chkQuestion15NA.Checked = False

        Me.chkQuestion16Yes.Checked = False
        Me.chkQuestion16No.Checked = False
        Me.chkQuestion16NA.Checked = False

        Me.chkQuestion17Yes.Checked = False
        Me.chkQuestion17No.Checked = False
        Me.chkQuestion17NA.Checked = False

        Me.chkCommissionNo.Checked = False
        Me.chkCommissionYes.Checked = False

        Me.chkforCommission.Checked = False
        Me.chkforHeadofOPC.Checked = False

        Me.chkOPCHeadNo.Checked = False
        Me.chkOPCHeadYes.Checked = False

        Me.chkPMHeadUndecided.Checked = False
        Me.chkPMHeadYes.Checked = False

        Me.chkUSTChiefNo.Checked = False
        Me.chkUSTChiefUndecided.Checked = False
        Me.chkUSTChiefYes.Checked = False

        Me.chkQuestion3Yes.Checked = False
        Me.chkQuestion3No.Checked = False
        Me.chkQuestion3NA.Checked = False

        Me.chkQuestion4Yes.Checked = False
        Me.chkQuestion4No.Checked = False
        Me.chkQuestion4NA.Checked = False

        Me.chkQuestion7Yes.Checked = False
        Me.chkQuestion7No.Checked = False
        Me.chkQuestion7NA.Checked = False

        Me.chkQuestion9Yes.Checked = False
        Me.chkQuestion9No.Checked = False
        Me.chkQuestion9NA.Checked = False

        Me.chkQuestion11Yes.Checked = False
        Me.chkQuestion11No.Checked = False
        Me.chkQuestion11NA.Checked = False

        Me.chkQuestion12Yes.Checked = False
        Me.chkQuestion12No.Checked = False
        Me.chkQuestion12NA.Checked = False

        Me.chkQuestion13Yes.Checked = False
        Me.chkQuestion13No.Checked = False
        Me.chkQuestion13NA.Checked = False

    End Sub


    Private Sub ModifyActivity()
        Dim frmAddActivity As Activity
        Dim MyFrm As MusterContainer
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand

        Try
            MyFrm = MdiParent
            If ugActivitiesandDocuments.ActiveRow.Band.Index = 0 Then

                frmAddActivity = New Activity(oLustEvent)
                frmAddActivity.CallingForm = Me
                frmAddActivity.Mode = 1 ' Modify
                frmAddActivity.EventActivityID = ugActivitiesandDocuments.ActiveRow.Cells("Event_Activity_ID").Value
                frmAddActivity.TFStatus = oLustEvent.MGPTFStatus

                frmAddActivity.ShowDialog()

                GetLustActivitiesForEvent(oLustEvent.ID)
                GetLustRemediationForEvent(oLustEvent.ID)

                MyFrm.RefreshCalendarInfo()
                MyFrm.LoadDueToMeCalendar()
                MyFrm.LoadToDoCalendar()
            Else
                MsgBox("You must select an Activity to modify.")
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Modify Activity" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmAddActivity = Nothing

        End Try
    End Sub

    Private Sub ModifyDocument()
        Dim frmAddDocument As New Document(oLustEvent)
        Dim tmpDate As Date
        Dim MyFrm As MusterContainer

        Try
            MyFrm = MdiParent
            If ugActivitiesandDocuments.ActiveRow.Band.Index = 1 Then

                frmAddDocument.CallingForm = Me
                frmAddDocument.Mode = 1 ' Add
                frmAddDocument.EventActivityID = ugActivitiesandDocuments.ActiveRow.ParentRow.Cells("Event_Activity_ID").Value
                frmAddDocument.EventDocumentID = ugActivitiesandDocuments.ActiveRow.Cells("EVENT_ACTIVITY_DOCUMENT_ID").Value
                frmAddDocument.TFStatus = oLustEvent.MGPTFStatus
                Dim user As BusinessLogic.pUser = MusterContainer.AppUser
                frmAddDocument.IsTech = IIf(user.DefaultModule = UIUtilsGen.ModuleID.Financial, False, True)

                frmAddDocument.ActivityName = ugActivitiesandDocuments.ActiveRow.ParentRow.Cells("Activity").Value
                frmAddDocument.StartDate = IIf(ugActivitiesandDocuments.ActiveRow.ParentRow.Cells("Start_Date").Value Is DBNull.Value, tmpDate, ugActivitiesandDocuments.ActiveRow.ParentRow.Cells("Start_Date").Value)

                frmAddDocument.ShowDialog()

                UpdateEventDates()
                GetLustActivitiesForEvent(oLustEvent.ID)
                GetLustRemediationForEvent(oLustEvent.ID)
                MyFrm.RefreshCalendarInfo()
                MyFrm.LoadDueToMeCalendar()
                MyFrm.LoadToDoCalendar()
            Else
                MsgBox("You must select a Document to modify.")
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Modify Document" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmAddDocument = Nothing

        End Try
    End Sub

    Private Sub ToggleActivitiesGrid()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        If chkShowallActivities.Checked = False Then
            For Each ugrow In ugActivitiesandDocuments.Rows
                If ugrow.Band.Index = 0 Then
                    If ugrow.Cells("Closed_Date").Value Is DBNull.Value Or ugrow.Cells("Activity").Text.StartsWith("NFA") Then
                        ugrow.Hidden = False
                    Else
                        ugrow.Hidden = True
                    End If
                End If
            Next
        Else
            For Each ugrow In ugActivitiesandDocuments.Rows
                If ugrow.Band.Index = 0 Then
                    ugrow.Hidden = False
                End If
            Next
        End If
        SetActiveActivitiesRow()

    End Sub

    Private Sub ToggleDocumentsGrid()
        Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        If chkShowallDocuments.Checked = False Then
            For Each ugrow In ugActivitiesandDocuments.Rows
                If ugrow.Band.Index = 0 Then
                    ChildBand = ugrow.ChildBands(0)
                    For Each Childrow In ChildBand.Rows
                        If Not (Childrow.Cells("Closed").Value Is DBNull.Value) Then
                            Childrow.Hidden = True
                        End If
                    Next
                End If
            Next
        Else
            For Each ugrow In ugActivitiesandDocuments.Rows
                ChildBand = ugrow.ChildBands(0)
                For Each Childrow In ChildBand.Rows
                    Childrow.Hidden = False
                Next
            Next
        End If
        SetActiveActivitiesRow()
    End Sub


    Private Sub SetActivitiesGrid()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            For Each ugrow In ugActivitiesandDocuments.Rows
                If ugrow.Band.Index = 0 Then

                    ' Per J.  No grid editing...
                    ugrow.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                    ugrow.Cells("Activity").Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    If Not (ugrow.Cells("Activity_Type_ID").Value = 675 _
                        Or ugrow.Cells("Activity_Type_ID").Value = 676 _
                        Or ugrow.Cells("Activity_Type_ID").Value = 691 _
                        Or ugrow.Cells("Activity_Type_ID").Value = 692 _
                        Or ugrow.Cells("Activity_Type_ID").Value = 694 _
                        Or ugrow.Cells("Activity_Type_ID").Value = 695 _
                        Or ugrow.Cells("Activity_Type_ID").Value = 1530) _
                    Then
                        ugrow.Cells("1ST_GWS_BELOW").Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ugrow.Cells("2ND_GWS_BELOW").Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ugrow.Cells("1ST_GWS_BELOW").Appearance.BackColor = Color.LightGray
                        ugrow.Cells("2ND_GWS_BELOW").Appearance.BackColor = Color.LightGray
                    End If

                    If (ugrow.Cells("DaysSinceStart").Value > ugrow.Cells("Act_By").Value) And (ugrow.Cells("Closed_Date").Value Is DBNull.Value And ugrow.Cells("Tech_Completed_Date").Value Is DBNull.Value) Then
                        ugrow.CellAppearance.BackColor = Color.Red
                    ElseIf ugrow.Cells("DaysSinceStart").Value > ugrow.Cells("Warn").Value And (ugrow.Cells("Closed_Date").Value Is DBNull.Value And ugrow.Cells("Tech_Completed_Date").Value Is DBNull.Value) Then
                        ugrow.CellAppearance.BackColor = Color.Yellow
                    End If
                    ChildBand = ugrow.ChildBands(0)
                    For Each Childrow In ChildBand.Rows
                        ' Per J.  No grid editing...
                        Childrow.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                        Childrow.Cells("Document").Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        If Not Childrow.Cells("Due").Value Is DBNull.Value Then
                            If Childrow.Cells("Due").Value < Now.Date And (Childrow.Cells("Closed").Value Is DBNull.Value And Childrow.Cells("Received").Value Is DBNull.Value And Childrow.Cells("To Financial").Value Is DBNull.Value) Then
                                If Childrow.Cells("Extension").Value Is DBNull.Value Then
                                    Childrow.CellAppearance.BackColor = Color.Orange
                                ElseIf Childrow.Cells("Extension").Value < Now.Date Then
                                    Childrow.CellAppearance.BackColor = Color.Orange
                                End If
                            End If
                        End If
                        If Not (Childrow.Cells("RevisionDate").Value Is DBNull.Value) Then
                            Childrow.Cells("Due").Appearance.BackColor = Color.Yellow
                        End If

                        ' #2511
                        If Not Childrow.Cells("Doc_Type").Value Is Nothing Then
                            If Childrow.Cells("Doc_Type").Value = 919 Then
                                Childrow.Cells("Received").Appearance.BackColor = Color.LightGray
                            End If
                            If Childrow.Cells("Doc_Type").Value = 919 Or Childrow.Cells("Doc_Type").Value = 918 Then
                                Childrow.Cells("To Financial").Appearance.BackColor = Color.LightGray
                                Childrow.Cells("Extension").Appearance.BackColor = Color.LightGray
                                Childrow.Cells("RevisionDate").Appearance.BackColor = Color.LightGray
                            Else
                                If oLustEvent.MGPTFStatus = 617 Or oLustEvent.MGPTFStatus = 620 Then
                                    Childrow.Cells("To Financial").Appearance.BackColor = Color.LightGray
                                Else
                                    Childrow.Cells("To Financial").Appearance.BackColor = Childrow.CellAppearance.BackColor
                                End If
                                Childrow.Cells("Extension").Appearance.BackColor = Childrow.CellAppearance.BackColor
                                Childrow.Cells("RevisionDate").Appearance.BackColor = Childrow.CellAppearance.BackColor
                            End If
                        Else
                            Childrow.Cells("Received").Appearance.BackColor = Childrow.CellAppearance.BackColor
                            If oLustEvent.MGPTFStatus = 617 Or oLustEvent.MGPTFStatus = 620 Then
                                Childrow.Cells("To Financial").Appearance.BackColor = Color.LightGray
                            Else
                                Childrow.Cells("To Financial").Appearance.BackColor = Childrow.CellAppearance.BackColor
                            End If
                            Childrow.Cells("Extension").Appearance.BackColor = Childrow.CellAppearance.BackColor
                            Childrow.Cells("RevisionDate").Appearance.BackColor = Childrow.CellAppearance.BackColor
                        End If
                        If Childrow.Cells("Document").Text.IndexOf("MDEQ SOW") > -1 Then
                            If Childrow.Cells("Due").Appearance.BackColor.ToString <> Color.Yellow.ToString Then
                                Childrow.Cells("Due").Appearance.BackColor = Color.LightGray
                            End If
                            Childrow.Cells("Extension").Appearance.BackColor = Color.LightGray
                            Childrow.Cells("RevisionDate").Appearance.BackColor = Color.LightGray
                            Childrow.Cells("Received").Appearance.BackColor = Color.LightGray
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Setting Activities Grid" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub

    Private Sub SetActiveActivitiesRow()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim bolOneActive As Boolean = False

        If ugActivitiesandDocuments.Rows.Count > 0 Then
            For Each ugrow In ugActivitiesandDocuments.Rows
                If ugrow.Band.Index = 0 Then
                    If ugrow.Hidden = False Then
                        ugrow.Activated = True
                        Exit For
                    End If
                End If
            Next
        End If

    End Sub
    Private Sub ProcessTankAndPipe()
        Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim strTanksAndPipes As String

        strTanksAndPipes = ""

        For Each ugrow In ugTankandPipes.Rows
            If ugrow.Band.Index = 0 Then
                If ugrow.Cells("Included").Value = True Then
                    If strTanksAndPipes.Length > 0 Then
                        strTanksAndPipes &= "|"
                    End If
                    strTanksAndPipes &= "T" & ugrow.Cells("Tank ID").Value
                End If
            End If
        Next

        For Each ugrow In ugTankandPipes.Rows
            If ugrow.Band.Index = 0 Then
                ChildBand = ugrow.ChildBands(0)
                For Each Childrow In ChildBand.Rows
                    If Childrow.Cells("Included").Value = True Then
                        If strTanksAndPipes.Length > 0 Then
                            strTanksAndPipes &= "|"
                        End If
                        strTanksAndPipes &= "P" & Childrow.Cells("Pipe_ID").Value
                    End If
                Next
            End If
        Next
        If oLustEvent.TankandPipe <> strTanksAndPipes Then
            oLustEvent.TankandPipe = strTanksAndPipes
            EvaluateTFChecklist(True)
        End If

        TestForConfirmedSpill()
    End Sub

    Private Sub ModifyRemediation()
        Dim frmRemediation As New RemediationSystem
        Try
            If ugActivitiesandDocuments.ActiveRow.Band.Index = 0 Then
                frmRemediation.CallingForm = Me
                frmRemediation.Mode = 1 ' Modify
                frmRemediation.nActivityID = ugRemediationSystem.ActiveRow.Cells("Event_Activity_ID").Value
                frmRemediation.nSystemID = ugRemediationSystem.ActiveRow.Cells("REM_SYSTEM_ID").Value
                frmRemediation.ShowDialog()

                GetLustRemediationForEvent(oLustEvent.ID)
            Else
                MsgBox("You must select an Remediation System to modify.")
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Processing Modify Remediation" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmRemediation = Nothing

        End Try
    End Sub


    Private Sub ProcessSaveEvent(Optional ByVal passChecklist As Boolean = False)
        Dim strMessage As String
        Dim bolEligibityChanged As Boolean = False
        Dim oCalendar As New MUSTER.BusinessLogic.pCalendar
        Dim MyFrm As MusterContainer

        Try
            MyFrm = MdiParent

            If oLustEvent.ReportDate = "01/01/0001" Then
                MsgBox("Date of Report is required")
                dtDateofReport.Focus()
                Exit Sub
            End If
            bolEligibityChanged = oLustEvent.IsEligibityDirty
            If Me.bolMGPTFCheckList AndAlso Not passChecklist Then
                If MsgBox("Do you wish to re-evaluate automated checklist questions?", MsgBoxStyle.YesNo, "Re-evaluate questions") = MsgBoxResult.Yes Then
                    EvaluateTFChecklist(True)
                Else
                    EvaluateTFChecklist(False)
                End If
                bolMGPTFCheckList = False
            End If

            ' #470 
            ' If any of the trust fund eligibility questions are answered yes/no or mgptf status = (stfs or stfs-direct)
            ' and there is an open TF Paperwork activity
            ' prompt user "Do you want to close the TF Paperwork activity"
            If chkPMHeadYes.Checked Or _
                chkUSTChiefYes.Checked Or chkUSTChiefNo.Checked Or _
                chkOPCHeadYes.Checked Or chkOPCHeadNo.Checked Or _
                chkCommissionYes.Checked Or chkCommissionNo.Checked Or _
                oLustEvent.MGPTFStatus = 618 Or oLustEvent.MGPTFStatus = 619 Then
                For Each ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugActivitiesandDocuments.Rows
                    If ugrow.Cells("Activity").Text.IndexOf("Determining Trust Fund Eligibility") > -1 And ugrow.Cells("Closed_Date").Value Is DBNull.Value Then
                        Dim msgResult As MsgBoxResult = MsgBox("There is an open 'Determining Trust Fund Eligibility' activity." + vbCrLf + _
                                                                "Do you want to close the activity" + vbCrLf + vbCrLf + _
                                                                "Yes - Close the Activity (Closed Date = Today)" + vbCrLf + _
                                                                "No - Leave the Activity Open and continue Save Event" + vbCrLf + _
                                                                "Cancel - Abort Save Event", MsgBoxStyle.YesNoCancel, "Open TF Paperwork Exists")
                        If msgResult = MsgBoxResult.Yes Then
                            Dim oLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
                            oLustActivity.Retrieve(ugrow.Cells("Event_Activity_ID").Value)
                            oLustActivity.Closed = Today.Date
                            oLustActivity.Save(UIUtilsGen.ModuleID.Technical, MusterContainer.AppUser.UserKey, returnVal)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If

                            ugrow.Cells("Closed_Date").Value = Today.Date
                            ugrow.Cells("Tech_Completed_Date").Value = Today.Date
                        ElseIf msgResult = MsgBoxResult.Cancel Then
                            Exit Sub
                        End If
                        Exit For
                    End If
                Next
            End If

            If oLustEvent.ID <= 0 Then
                oLustEvent.CreatedBy = MusterContainer.AppUser.ID
            Else
                oLustEvent.ModifiedBy = MusterContainer.AppUser.ID
            End If

            TestForConfirmedSpill()
            If lblEventID.Text = "Add Event" Then
                'Create TF Checklist
                If Not passChecklist Then
                    EvaluateTFChecklist()
                End If


                'Save Lust Event
                oLustEvent.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                nSavedEventID = oLustEvent.ID
                nLastEventID = oLustEvent.ID
                nCurrentEventID = oLustEvent.ID
                'Create New Site Activity
                GenerateActivity_NewSite()

                'Create Flag
                GenerateFlag_OnAdd()

                strMessage = "Lust Event Successfully Added"

                SetupAfterAdd()
            Else


                'Update Record
                oLustEvent.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                nSavedEventID = oLustEvent.ID

                ' #2188
                If oLustEvent.Priority > 0 Then
                    cmbPriority.SelectedValue = oLustEvent.Priority
                Else
                    cmbPriority.SelectedIndex = -1
                    If cmbPriority.SelectedIndex <> -1 Then
                        cmbPriority.SelectedIndex = -1
                    End If
                End If

                If bolEligibityChanged And EligibityIsChecked() And Not EligibityIsComplete() Then
                    oLetter.GenerateTFCheckList(oLustEvent.ID, oLustEvent.FacilityID, oOwner, False, oLustEvent.EVENTSEQUENCE)
                    If EligibityIsComplete() Then
                        oCalendar.Retrieve(oLustEvent.GetSentToPMCalendarID(Me.PMHead_UserID, oLustEvent.ID))
                        If Not IsNothing(oCalendar) Then
                            oCalendar.Completed = True
                            oCalendar.Save()
                        End If
                    End If
                End If

                If oLustEvent.PM <> nOriginalPM And nOriginalPM <> 0 Then
                    oLustEvent.CalendarEntries(cmbProjectManager.Text, "Transferred Lust Event for ID: " & oLustEvent.FacilityID & " Event " & oLustEvent.EVENTSEQUENCE, True, False, String.Empty, UserID, Now.Date, Now.Date, oLustEvent.ID)


                    'moves to due list to new PM
                    Dim oldCal As New BusinessLogic.pCalendar

                    Dim ds As DataSet

                    'Seeks any available records in calendars for luste event of facility 
                    ds = oldCal.GetDataSet(String.Format("spGetlustcalendarBySiteID {0}", oLustEvent.FacilityID))


                    'if found, create instances of calendar to update 
                    If Not ds Is Nothing AndAlso Not ds.Tables Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                        'each calendar item per row
                        For Each row As DataRow In ds.Tables(0).Rows
                            oldCal.Retrieve(row("CALENDAR_INFO_ID"))

                            'change PM
                            Dim user As New BusinessLogic.pUser
                            user.Retrieve(oLustEvent.PM)
                            oldCal.UserId = user.ID
                            user = Nothing

                            'send update
                            oldCal.Save()

                        Next

                    End If

                    oldCal = Nothing


                End If



                If Not MyFrm Is Nothing Then
                    MyFrm.RefreshCalendarInfo()
                    MyFrm.LoadDueToMeCalendar()
                    MyFrm.LoadToDoCalendar()
                End If
                strMessage = "Lust Event Successfully Updated"
                SetupModifyLustEventForm(True, False)
            End If

            If bolNewRelease Then
                'Create Written Confirmation Report Letter
                oLetter.GenerateTechLetter(oLustEvent.FacilityID, "Written Confirmation Of Release", "ConfirmRel", "Written Confirmation Of Release", "Written_Confirmation_of_a_Release_Letter_Template1.doc", DateAdd(DateInterval.Day, 30, Now()), oLustEvent.ID, oOwner, 0, oLustEvent.EVENTSEQUENCE, UIUtilsGen.EntityTypes.LUST_Event)
                ShowDocumentList()
                bolNewRelease = False
            End If

            GetLustEventsForFacility()
            UIUtilsGen.PopulateOwnerFacilities(oOwner, Me, Integer.Parse(Me.lblOwnerIDValue.Text))
            MsgBox(strMessage)

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Processing Save Event" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally

        End Try
    End Sub
    Private Sub UpdateEventDates()

        UIUtilsGen.SetDatePickerValue(dtLastLDR, oLustEvent.LastLDR)
        UIUtilsGen.SetDatePickerValue(dtLastPTT, oLustEvent.LastPTT)
        UIUtilsGen.SetDatePickerValue(dtLastGWS, oLustEvent.LastGWS)

    End Sub
#End Region

#End Region

#Region "TAB Operations"
    Private Sub tbCntrlTechnical_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlTechnical.Click
        Dim MyFrm As MusterContainer
        Dim nFacilityID As Integer
        '  TO DO - Elango on Dec 28 2004 
        Try
            Select Case tbCntrlTechnical.SelectedTab.Name
                Case tbPageFacilityDetail.Name
                    nCurrentEventID = -1
                    If ugFacilityList.Rows.Count <> 0 Then
                        If lblFacilityIDValue.Text = String.Empty Then
                            bolTabClick = True
                            ugOwnerDetailsFacilities_DoubleClick(sender, e)
                            bolTabClick = False
                        Else
                            MyFrm = Me.MdiParent
                            MyFrm.FlagsChanged(Me.lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Technical", Me.Text)
                        End If
                        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Technical")
                    End If

                    If ugFacilityList.Rows.Count <= 0 And Me.lblOwnerIDValue.Text <> String.Empty Then
                        Dim msgResult As MsgBoxResult
                        msgResult = MsgBox("No facilities found for Owner " + lblOwnerIDValue.Text)
                        tbCntrlTechnical.SelectedTab = tbPageOwnerDetail
                        Exit Sub
                    End If
                    Me.Text = "Technical - Facility Detail - " & lblFacilityIDValue.Text & " (" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"

                    If oOwner.Facility.ID > 0 Then
                        If Not tbCntrlFacility.TabPages.Contains(tbPageOwnerContactList) Then
                            tbCntrlFacility.TabPages.Add(tbPageOwnerContactList)
                            lblOwnerContacts.Text = "Facility Contacts"
                        End If
                        'LoadContacts(ugOwnerContacts, oOwner.Facility.ID, 6)
                    End If
                    If ugFacilityList.Rows.Count > 0 And Not oOwner.Facilities.ID > 0 Then
                        If ugFacilityList.ActiveRow Is Nothing Then
                            ugFacilityList.ActiveRow = ugFacilityList.Rows(0)
                        End If
                        nFacilityID = ugFacilityList.ActiveRow.Cells("FacilityID").Value
                        Me.PopFacility(nFacilityID)
                    Else
                        Me.PopFacility(oOwner.Facilities.ID)
                    End If
                Case tbPageLUSTEvent.Name
                    tbPageLUSTEvent.Enabled = True
                    If dgLUSTEvents.Rows.Count <> 0 Then
                        If dgLUSTEvents.ActiveRow Is Nothing Then
                            dgLUSTEvents.ActiveRow = dgLUSTEvents.Rows(0)
                        End If
                        If (nLastEventID <> dgLUSTEvents.ActiveRow.Cells("EVENT_ID").Text) And nLastEventID >= 0 Then
                            SetupModifyLustEventForm()
                        Else
                            If nLastEventID <> 0 Then
                                SetupModifyLustEventForm()
                            End If
                            LoadContacts(ugERACContacts, oLustEvent.ID, 7)
                        End If
                    End If
                    If dgLUSTEvents.Rows.Count <= 0 And lblEventID.Text = "" Then
                        If lblFacilityIDValue.Text <> String.Empty Then
                            If lblEventInfoDisplay.Text = "-" Then
                                lblEventInfoDisplay_Click(sender, e)
                            End If
                            If lblReleaseInfoDisplay.Text = "-" Then
                                lblReleaseInfoDisplay_Click(sender, e)
                            End If
                            If lblActivitiesDocumentsDisplay.Text = "-" Then
                                lblActivitiesDocumentsDisplay_Click(sender, e)
                            End If
                            If lblFundsEligibilityDisplay.Text = "-" Then
                                lblFundsEligibilityDisplay_Click(sender, e)
                            End If
                            If lblCommentsDisplay.Text = "-" Then
                                lblCommentsDisplay_Click(sender, e)
                            End If
                            If lblRemediationSystemsDisplay.Text = "-" Then
                                lblRemediationSystemsDisplay_Click(sender, e)
                            End If
                            If lblERACandContactsDisplay.Text = "-" Then
                                lblERACandContactsDisplay_Click(sender, e)
                            End If
                            'If lblPermitsDisplay.Text = "-" Then
                            '    lblPermitsDisplay_Click(sender, e)
                            'End If
                            tbPageLUSTEvent.Enabled = False
                            MsgBox("No events found for facility:  " + lblFacilityIDValue.Text)
                        Else
                            If lblEventInfoDisplay.Text = "-" Then
                                lblEventInfoDisplay_Click(sender, e)
                            End If
                            If lblReleaseInfoDisplay.Text = "-" Then
                                lblReleaseInfoDisplay_Click(sender, e)
                            End If
                            If lblActivitiesDocumentsDisplay.Text = "-" Then
                                lblActivitiesDocumentsDisplay_Click(sender, e)
                            End If
                            If lblFundsEligibilityDisplay.Text = "-" Then
                                lblFundsEligibilityDisplay_Click(sender, e)
                            End If
                            If lblCommentsDisplay.Text = "-" Then
                                lblCommentsDisplay_Click(sender, e)
                            End If
                            If lblRemediationSystemsDisplay.Text = "-" Then
                                lblRemediationSystemsDisplay_Click(sender, e)
                            End If
                            If lblERACandContactsDisplay.Text = "-" Then
                                lblERACandContactsDisplay_Click(sender, e)
                            End If
                            'If lblPermitsDisplay.Text = "-" Then
                            '    lblPermitsDisplay_Click(sender, e)
                            'End If
                            tbPageLUSTEvent.Enabled = False
                            MsgBox("No facility selected")

                        End If

                    End If
                    Me.Text = "Technical LUST Events - Facility #" & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
                    nSavedEventID = 0

                Case tbPageOwnerDetail.Name
                    nCurrentEventID = -1
                    Me.Text = "Technical - Owner Detail (" & txtOwnerName.Text & ")"
                    'If lblOwnerIDValue.Text <> String.Empty Then
                    '    ' Me.PopulateOwnerFacilities()
                    '    UIUtilsGen.PopulateOwnerFacilities(oOwner, Me, Integer.Parse(Me.lblOwnerIDValue.Text))
                    'End If
                    PopulateOwnerInfo(oOwner.ID)
                    If oOwner.ID > 0 Then
                        If Not tbCtrlOwner.TabPages.Contains(tbPageOwnerContactList) Then
                            tbCtrlOwner.TabPages.Add(tbPageOwnerContactList)
                            lblOwnerContacts.Text = "Owner Contacts"
                        End If
                        'LoadContacts(ugOwnerContacts, oOwner.ID, 9)
                        MyFrm = Me.MdiParent
                        MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, UIUtilsGen.EntityTypes.Owner, "Technical", Me.Text)
                    End If
                Case tbPageOwnerFacilities.Name
                    nCurrentEventID = -1
                Case tbPageSummary.Name
                    nCurrentEventID = -1
                    Me.Text = "Technical - Owner Summary (" & txtOwnerName.Text & ")"
                    UIUtilsGen.PopulateOwnerSummary(oOwner, Me)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    'Private Sub tbCtrlOwner_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlOwner.Click
    '    'Select Case tbCtrlOwner.SelectedTab.Name.ToUpper
    '    '    Case "tbPrevFacs".ToUpper
    '    '        tbCtrlOwner.SelectedTab = tbPageOwnerFacilities
    '    '    Case "tbPageOwnerFacilities".ToUpper
    '    '        If tbCtrlOwner.Contains(Me.tbPrevFacs) Then
    '    '            Me.tbCtrlOwner.TabPages.RemoveAt(1)
    '    '        End If
    '    'End Select
    'End Sub
#End Region

#Region "Event Handlers for Expanding/Collapsing the different Sections of the Lust Event form"


    Private Sub ShowHideControl(ByVal ObjControl As Control)
        If ObjControl.Visible Then
            ObjControl.Visible = False
        Else
            ObjControl.Visible = True
        End If
    End Sub


    Private Sub lblEventInfoDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblEventInfoDisplay.Click, lblEventInfoHead.Click
        If lblEventInfoDisplay.Text = "+" Then
            lblEventInfoDisplay.Text = "-"
        Else
            lblEventInfoDisplay.Text = "+"
        End If
        ShowHideControl(PnlEventInfoDetails)
    End Sub


    Private Sub lblReleaseInfoDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblReleaseInfoDisplay.Click, lblReleaseInfoHead.Click
        If lblReleaseInfoDisplay.Text = "+" Then
            lblReleaseInfoDisplay.Text = "-"
        Else
            lblReleaseInfoDisplay.Text = "+"
        End If
        ShowHideControl(PnlReleaseInfoDetails)
    End Sub

    Private Sub lblFundsEligibilityDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblFundsEligibilityDisplay.Click, lblFundsEligibilityHead.Click
        If lblFundsEligibilityDisplay.Text = "+" Then
            lblFundsEligibilityDisplay.Text = "-"
        Else
            lblFundsEligibilityDisplay.Text = "+"
        End If
        ShowHideControl(pnlFundsEligibilityDetails)
    End Sub

    Private Sub lblActivitiesDocumentsDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblActivitiesDocumentsDisplay.Click, lblActivitiesDocumentsHead.Click
        If lblActivitiesDocumentsDisplay.Text = "+" Then
            lblActivitiesDocumentsDisplay.Text = "-"
        Else
            lblActivitiesDocumentsDisplay.Text = "+"
        End If

        ShowHideControl(pnlActivitiesDocumentsDetails)
    End Sub

    Private Sub lblCommentsDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCommentsDisplay.Click, lblCommentsHead.Click
        If lblCommentsDisplay.Text = "+" Then
            lblCommentsDisplay.Text = "-"
        Else
            lblCommentsDisplay.Text = "+"
        End If

        ShowHideControl(pnlCommentsDetails)
    End Sub

    Private Sub lblRemediationSystemsDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblRemediationSystemsDisplay.Click, lblRemediationSystemsHead.Click
        If lblRemediationSystemsDisplay.Text = "+" Then
            lblRemediationSystemsDisplay.Text = "-"
        Else
            lblRemediationSystemsDisplay.Text = "+"
        End If

        ShowHideControl(Me.pnlRemediationSystemsDetails)
    End Sub

    Private Sub lblERACandContactsDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblERACandContactsDisplay.Click, lblERACandContactsHead.Click
        If lblERACandContactsDisplay.Text = "+" Then
            lblERACandContactsDisplay.Text = "-"
        Else
            lblERACandContactsDisplay.Text = "+"
        End If

        ShowHideControl(Me.pnlERACandContactsDetails)
    End Sub


    'Private Sub lblPermitsDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If lblPermitsDisplay.Text = "+" Then
    '        lblPermitsDisplay.Text = "-"
    '    Else
    '        lblPermitsDisplay.Text = "+"
    '    End If

    '    ShowHideControl(Me.pnlPermitsDetails)
    'End Sub
#End Region

#Region "External Events"
    Public Sub FacilitiesChanged(ByVal bolstate As Boolean) Handles oOwner.evtFacilitiesChanged
        SetFacilitySaveCancel(bolstate)
    End Sub
    Public Sub FacilityChanged(ByVal bolstate As Boolean) Handles oOwner.evtFacilityChanged
        SetFacilitySaveCancel(bolstate)
    End Sub
    Private Sub SetFacilitySaveCancel(ByVal bolState As Boolean)
        Me.btnFacilitySave.Enabled = bolState
        btnFacilityCancel.Enabled = bolState
    End Sub
#End Region

#Region "Lookup Functions"

    Private Sub PopulateLustEventLookups()
        PopulateLustEventStatus()
        PopulateLustMGPTFStatus()
        PopulateLustLeakPriority()
        PopulateLustReleaseStatus()
        PopulateLustReleaseIdentifiedBy()
        PopulateLustReleaseLocation()
        PopulateLustReleaseExtent()
        PopulateLustCause()
        PopulateLustProjectManager()
        PopulateLustSuspectedSource()
        'PopulateERACCompanyName()
        'PopulateIRACCompanyName()
    End Sub

    Private Sub PopulateLustEventStatus()
        'If bolLoading Then Exit Sub
        Try
            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustEventStatus
            If Not IsNothing(dtLustEventStatus) Then
                cmbEventStatus.DataSource = dtLustEventStatus
                cmbEventStatus.DisplayMember = "PROPERTY_NAME"
                cmbEventStatus.ValueMember = "PROPERTY_ID"
            Else
                cmbEventStatus.DataSource = Nothing
            End If
            If lblEventID.Text <> "Add Event" Then
                Me.cmbEventStatus.SelectedIndex = -1
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Event Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub PopulateLustMGPTFStatus()
        'If bolLoading Then Exit Sub
        Try
            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustMGPTFStatus
            If Not IsNothing(dtLustEventStatus) Then

                cmbMGPTFStatus.DataSource = dtLustEventStatus
                cmbMGPTFStatus.DisplayMember = "PROPERTY_NAME"
                cmbMGPTFStatus.ValueMember = "PROPERTY_ID"
            Else
                cmbMGPTFStatus.DataSource = Nothing
            End If
            If lblEventID.Text <> "Add Event" Then
                cmbMGPTFStatus.SelectedIndex = -1
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust MGPTF Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub PopulateLustReleaseStatus()
        'If bolLoading Then Exit Sub
        Try
            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustReleaseStatus
            If Not IsNothing(dtLustEventStatus) Then

                cmbReleaseStatus.DataSource = dtLustEventStatus
                cmbReleaseStatus.DisplayMember = "PROPERTY_NAME"
                cmbReleaseStatus.ValueMember = "PROPERTY_ID"
            Else
                cmbReleaseStatus.DataSource = Nothing
            End If
            If lblEventID.Text <> "Add Event" Then
                cmbReleaseStatus.SelectedIndex = -1
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Release Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub PopulateLustLeakPriority()
        'If bolLoading Then Exit Sub
        Try
            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustLeakPriority
            If Not IsNothing(dtLustEventStatus) Then

                cmbPriority.DataSource = dtLustEventStatus
                cmbPriority.DisplayMember = "PROPERTY_NAME"
                cmbPriority.ValueMember = "PROPERTY_ID"
            Else
                cmbPriority.DataSource = Nothing
            End If
            cmbPriority.SelectedIndex = -1

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Leak Priority" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub PopulateLustReleaseIdentifiedBy()
        'If bolLoading Then Exit Sub
        Try
            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustIdentifiedBy
            If Not IsNothing(dtLustEventStatus) Then

                cmbIdentifiedBy.DataSource = dtLustEventStatus
                cmbIdentifiedBy.DisplayMember = "PROPERTY_NAME"
                cmbIdentifiedBy.ValueMember = "PROPERTY_ID"
            Else
                cmbIdentifiedBy.DataSource = Nothing
            End If
            cmbIdentifiedBy.SelectedValue = 0

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Release Identified By" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub



    Private Sub PopulateLustReleaseLocation()
        'If bolLoading Then Exit Sub
        Try
            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustReleaseLocation
            If Not IsNothing(dtLustEventStatus) Then

                cmbLocation.DataSource = dtLustEventStatus
                cmbLocation.DisplayMember = "PROPERTY_NAME"
                cmbLocation.ValueMember = "PROPERTY_ID"
            Else
                cmbLocation.DataSource = Nothing
            End If
            cmbLocation.SelectedValue = 0

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Release Location" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub PopulateLustReleaseExtent()
        'If bolLoading Then Exit Sub
        Try
            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustReleaseExtent
            If Not IsNothing(dtLustEventStatus) Then

                cmbExtent.DataSource = dtLustEventStatus
                cmbExtent.DisplayMember = "PROPERTY_NAME"
                cmbExtent.ValueMember = "PROPERTY_ID"
            Else
                cmbExtent.DataSource = Nothing
            End If
            cmbExtent.SelectedValue = 0

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Release Extent" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub PopulateLustCause()
        Try
            Dim dtLustCause As DataTable = oLustEvent.PopulateLustCause
            If Not IsNothing(dtLustCause) Then
                cmbCause.DataSource = dtLustCause
                cmbCause.DisplayMember = "PROPERTY_NAME"
                cmbCause.ValueMember = "PROPERTY_ID"
            Else
                cmbCause.DataSource = Nothing
            End If
            cmbCause.SelectedValue = 0
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Cause" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateLustProjectManager()
        'If bolLoading Then Exit Sub
        Try
            If IsNothing(PMHead_StaffID) Then
                Dim oUserInfo As New MUSTER.Info.UserInfo
                Dim oUser As New MUSTER.BusinessLogic.pUser

                oUserInfo = oUser.RetrievePMHead
                PMHead_UserID = oUserInfo.ID
                PMHead_StaffID = oUserInfo.UserKey

            End If
            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustProjectManager
            If Not IsNothing(dtLustEventStatus) Then

                cmbProjectManager.DataSource = dtLustEventStatus
                cmbProjectManager.DisplayMember = "USER_NAME"
                cmbProjectManager.ValueMember = "STAFF_ID"
            Else
                cmbProjectManager.DataSource = Nothing
            End If
            If IsNothing(PMHead_StaffID) Then
                PMHead_StaffID = -1
            End If

            cmbProjectManager.SelectedValue = PMHead_StaffID
            If PMHead_StaffID = -1 And cmbProjectManager.SelectedIndex <> -1 Then
                cmbProjectManager.SelectedIndex = -1
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Project Manager" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateLustSuspectedSource()
        'If bolLoading Then Exit Sub
        Dim rData As DataRow
        Try

            Dim dtLustEventStatus As DataTable = oLustEvent.PopulateLustSuspectedSource
            cmbSuspectedSource.DataSource = Nothing

            If Not IsNothing(dtLustEventStatus) Then

                cmbSuspectedSource.DataSource = dtLustEventStatus
                cmbSuspectedSource.DisplayMember = "PROPERTY_NAME"
                cmbSuspectedSource.ValueMember = "PROPERTY_ID"
            End If
            Me.cmbSuspectedSource.SelectedValue = 0

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Suspected Source" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub PopulateERACCompanyName()
    '    'If bolLoading Then Exit Sub
    '    Dim rData As DataRow
    '    Try

    '        Dim dtERACCompanyName As DataTable = oLustEvent.PopulateERACCompanyName
    '        cmbERAC.DataSource = Nothing

    '        If Not IsNothing(dtERACCompanyName) Then

    '            cmbERAC.DataSource = dtERACCompanyName
    '            cmbERAC.DisplayMember = "Company_Name"
    '            cmbERAC.ValueMember = "Company_ID"
    '        End If
    '        Me.cmbERAC.SelectedValue = 0

    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr = New ErrorReport(New Exception("Cannot Populate ERAC Company Name" + vbCrLf + ex.Message, ex))
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub PopulateIRACCompanyName()
    '    'If bolLoading Then Exit Sub
    '    Dim rData As DataRow
    '    Try

    '        Dim dtIRACCompanyName As DataTable = oLustEvent.PopulateIRACCompanyName
    '        cmbIRAC.DataSource = Nothing

    '        If Not IsNothing(dtIRACCompanyName) Then

    '            cmbIRAC.DataSource = dtIRACCompanyName
    '            cmbIRAC.DisplayMember = "Company_Name"
    '            cmbIRAC.ValueMember = "Company_ID"
    '        End If
    '        Me.cmbIRAC.SelectedValue = 0

    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr = New ErrorReport(New Exception("Cannot Populate IRAC Company Name" + vbCrLf + ex.Message, ex))
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region

#Region " (TF) Trust Fund Checklist "

    Private Sub btnViewCheckList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewCheckList.Click
        'Dim oLetter As New Reg_Letters
        Try
            If Me.bolMGPTFCheckList Then
                Dim boxResult As MsgBoxResult
                boxResult = MsgBox("Do you want to save event before viewing checklist?" + vbCrLf + _
                                    "Yes - Save and View Updated Checklist Info" + vbCrLf + _
                                    "No - View Previously Saved Checklist Info" + vbCrLf + _
                                    "Cancel - Abort View checklist", MsgBoxStyle.YesNoCancel, "Save Tech Event?")
                If boxResult = MsgBoxResult.Yes Then
                    btnSaveLUSTEvent.PerformClick()
                ElseIf boxResult = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If

            oLetter.GenerateTFCheckList(oLustEvent.ID, oLustEvent.FacilityID, oOwner, True, oLustEvent.EVENTSEQUENCE)

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot open the file in Word: " + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
        End Try
    End Sub

    Private Sub EvaluateTFChecklist(Optional ByVal bolRedoQuestions As Boolean = False)
        Dim strTFChecklist As String
        Dim arrayTF As Array
        Dim i As Int16

        If oLustEvent.TFCheckList Is Nothing Then
            strTFChecklist = "X|X|X|X|X|X|X|X|X|X|X|X|X|X|X|X|X"
        Else
            strTFChecklist = oLustEvent.TFCheckList
        End If

        If strTFChecklist.Length < 1 Then
            strTFChecklist = "X|X|X|X|X|X|X|X|X|X|X|X|X|X|X|X|X"
        End If
        arrayTF = Split(strTFChecklist, "|")

        If Not bolRedoQuestions Then

            If chkQuestion1Yes.Checked Then
                arrayTF(0) = "Y"
            ElseIf chkQuestion1No.Checked Then
                arrayTF(0) = "N"
            Else
                arrayTF(0) = "X"
            End If
            If chkQuestion2Yes.Checked Then
                arrayTF(1) = "Y"
            ElseIf chkQuestion2No.Checked Then
                arrayTF(1) = "N"
            Else
                arrayTF(1) = "X"
            End If

            If chkQuestion5Yes.Checked Then
                arrayTF(4) = "Y"
            ElseIf chkQuestion5No.Checked Then
                arrayTF(4) = "N"
            ElseIf Me.chkQuestion5NA.Checked Then
                arrayTF(4) = "A"
            Else
                arrayTF(4) = "X"
            End If

            If chkQuestion6Yes.Checked Then
                arrayTF(5) = "Y"
            ElseIf chkQuestion6No.Checked Then
                arrayTF(5) = "N"
            ElseIf Me.chkQuestion6NA.Checked Then
                arrayTF(5) = "A"
            Else
                arrayTF(5) = "X"
            End If

            If chkQuestion8Yes.Checked Then
                arrayTF(7) = "Y"
            ElseIf chkQuestion8No.Checked Then
                arrayTF(7) = "N"
            ElseIf Me.chkQuestion8NA.Checked Then
                arrayTF(7) = "A"
            Else
                arrayTF(7) = "X"
            End If

            If chkQuestion10Yes.Checked Then
                arrayTF(9) = "Y"
            ElseIf chkQuestion10No.Checked Then
                arrayTF(9) = "N"
            ElseIf Me.chkQuestion10NA.Checked Then
                arrayTF(9) = "A"
            Else
                arrayTF(9) = "X"
            End If

            If chkQuestion14Yes.Checked Then
                arrayTF(13) = "Y"
            ElseIf chkQuestion14No.Checked Then
                arrayTF(13) = "N"
            ElseIf Me.chkQuestion14NA.Checked Then
                arrayTF(13) = "A"
            Else
                arrayTF(13) = "X"
            End If

            If chkQuestion15Yes.Checked Then
                arrayTF(14) = "Y"
            ElseIf chkQuestion15No.Checked Then
                arrayTF(14) = "N"
            ElseIf Me.chkQuestion15NA.Checked Then
                arrayTF(14) = "A"
            Else
                arrayTF(14) = "X"
            End If

            If chkQuestion16Yes.Checked Then
                arrayTF(15) = "Y"
            ElseIf chkQuestion16No.Checked Then
                arrayTF(15) = "N"
            ElseIf Me.chkQuestion16NA.Checked Then
                arrayTF(15) = "A"
            Else
                arrayTF(15) = "X"
            End If

            If chkQuestion17Yes.Checked Then
                arrayTF(16) = "Y"
            ElseIf chkQuestion17No.Checked Then
                arrayTF(16) = "N"
            ElseIf Me.chkQuestion17NA.Checked Then
                arrayTF(16) = "A"
            Else
                arrayTF(16) = "X"
            End If
        End If


        If lblEventID.Text = "Add Event" Or bolRedoQuestions Then
            arrayTF(2) = GetTFAnswer3()
            arrayTF(3) = GetTFAnswer4()
            arrayTF(6) = GetTFAnswer7()
            arrayTF(8) = GetTFAnswer9()
            arrayTF(10) = GetTFAnswer11()
            arrayTF(11) = GetTFAnswer12()
            arrayTF(12) = GetTFAnswer13()
        End If

        strTFChecklist = ""
        For i = 0 To 16
            strTFChecklist &= arrayTF(i) & "|"
        Next

        strTFChecklist = Mid(strTFChecklist, 1, strTFChecklist.Length - 1)

        oLustEvent.TFCheckList = strTFChecklist

    End Sub

    Private Sub LoadTFChecklist()
        Dim strTFChecklist As String
        Dim strTFChecklist2 As String

        Dim arrayTF As Array

        strTFChecklist = oLustEvent.TFCheckList
        EvaluateTFChecklist(True)
        strTFChecklist2 = oLustEvent.TFCheckList

        If strTFChecklist2 <> strTFChecklist Then
            '     If Not strTFChecklist Is Nothing AndAlso strTFChecklist.Length > 5 Then
            '   MsgBox("Current Checklist values has changed due to changes in facility's tank/pipe status")
            ' End If

            oLustEvent.IsDirty = True
            strTFChecklist = strTFChecklist2
            oLustEvent.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
        End If



        If strTFChecklist.Length < 1 Then
            Exit Sub
        End If
        arrayTF = Split(strTFChecklist, "|")

        Select Case arrayTF(0)
            Case "Y"
                chkQuestion1Yes.Checked = True
            Case "N"
                chkQuestion1No.Checked = True
        End Select

        Select Case arrayTF(1)
            Case "Y"
                chkQuestion2Yes.Checked = True
            Case "N"
                chkQuestion2No.Checked = True
        End Select

        Select Case arrayTF(2)
            Case "Y"
                chkQuestion3Yes.Checked = True
            Case "N"
                chkQuestion3No.Checked = True
            Case "A"
                chkQuestion3NA.Checked = True
        End Select

        Select Case arrayTF(3)
            Case "Y"
                chkQuestion4Yes.Checked = True
            Case "N"
                chkQuestion4No.Checked = True
            Case "A"
                chkQuestion4NA.Checked = True
        End Select

        Select Case arrayTF(4)
            Case "Y"
                chkQuestion5Yes.Checked = True
            Case "N"
                chkQuestion5No.Checked = True
            Case "A"
                chkQuestion5NA.Checked = True
        End Select

        Select Case arrayTF(5)
            Case "Y"
                chkQuestion6Yes.Checked = True
            Case "N"
                chkQuestion6No.Checked = True
            Case "A"
                chkQuestion6NA.Checked = True
        End Select

        Select Case arrayTF(6)
            Case "Y"
                chkQuestion7Yes.Checked = True
            Case "N"
                chkQuestion7No.Checked = True
            Case "A"
                chkQuestion7NA.Checked = True
        End Select

        Select Case arrayTF(7)
            Case "Y"
                chkQuestion8Yes.Checked = True
            Case "N"
                chkQuestion8No.Checked = True
            Case "A"
                chkQuestion8NA.Checked = True
        End Select

        Select Case arrayTF(8)
            Case "Y"
                chkQuestion9Yes.Checked = True
            Case "N"
                chkQuestion9No.Checked = True
            Case "A"
                chkQuestion9NA.Checked = True
        End Select

        Select Case arrayTF(9)
            Case "Y"
                chkQuestion10Yes.Checked = True
            Case "N"
                chkQuestion10No.Checked = True
            Case "A"
                chkQuestion10NA.Checked = True
        End Select

        Select Case arrayTF(10)
            Case "Y"
                chkQuestion11Yes.Checked = True
            Case "N"
                chkQuestion11No.Checked = True
            Case "A"
                chkQuestion11NA.Checked = True
        End Select

        Select Case arrayTF(11)
            Case "Y"
                chkQuestion12Yes.Checked = True
            Case "N"
                chkQuestion12No.Checked = True
            Case "A"
                chkQuestion12NA.Checked = True
        End Select

        Select Case arrayTF(12)
            Case "Y"
                chkQuestion13Yes.Checked = True
            Case "N"
                chkQuestion13No.Checked = True
            Case "A"
                chkQuestion13NA.Checked = True
        End Select

        Select Case arrayTF(13)
            Case "Y"
                chkQuestion14Yes.Checked = True
            Case "N"
                chkQuestion14No.Checked = True
            Case "A"
                chkQuestion14NA.Checked = True
        End Select

        Select Case arrayTF(14)
            Case "Y"
                chkQuestion15Yes.Checked = True
            Case "N"
                chkQuestion15No.Checked = True
            Case "A"
                chkQuestion15NA.Checked = True
        End Select

        Select Case arrayTF(15)
            Case "Y"
                chkQuestion16Yes.Checked = True
            Case "N"
                chkQuestion16No.Checked = True
            Case "A"
                chkQuestion16NA.Checked = True
        End Select

        Select Case arrayTF(16)
            Case "Y"
                chkQuestion17Yes.Checked = True
            Case "N"
                chkQuestion17No.Checked = True
            Case "A"
                chkQuestion17NA.Checked = True
        End Select
    End Sub
    Private Function GetTFAnswer3() As String
        'Check the Notification form module TANK STATUS to determine if any of the tanks are 
        '“currently in use”.  If any tank is “currently in use” – mark “yes”.  If none of 
        'the tanks are “currently in use”, check the Notification module to determine the 
        'date the tanks were taken out of use.  If the date of any tank taken “out of use”
        'is greater than July 1, 1988, mark “yes”.  Otherwise – mark “no”
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        'Dim oTankInfo As MUSTER.Info.TankInfo
        Dim strResponse As String

        strResponse = "N"
        Try
            For Each ugrow In ugTankandPipes.Rows
                If ugrow.Band.Index = 0 Then
                    'If ugrow.Cells("Included").Value = True Then
                    If ugrow.Cells("Status").Text.IndexOf("Currently In Use") > -1 Then
                        strResponse = "Y"
                        Exit For
                    ElseIf Not ugrow.Cells("DATELASTUSED").Value Is DBNull.Value Then
                        If Date.Compare(ugrow.Cells("DATELASTUSED").Value, CDate("07/01/1988")) > 0 Then
                            strResponse = "Y"
                            Exit For
                        End If
                    End If
                    'For Each oTankInfo In oOwner.Facility.TankCollection.Values
                    '    If CInt(ugrow.Cells("TANK ID").Value) = oTankInfo.TankId Then
                    '        If oTankInfo.TankStatus = 424 Then
                    '            strResponse = "Y"
                    '        Else
                    '            If oTankInfo.DateLastUsed.Date > CDate("07/01/1988") Then
                    '                strResponse = "Y"
                    '            End If
                    '        End If
                    '    End If
                    'Next
                    'End If
                End If
            Next

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error with TF Trustfund Question 3:  " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            GetTFAnswer3 = strResponse

        End Try


    End Function
    Private Function GetTFAnswer4() As String
        'Query the Fee Module Running Balance to determine if any fees exist prior to the 
        'release date (release date is listed in the LUST Module).  If  fees exist – mark
        '“no”, otherwise mark “yes”
        ' if confirmed date is null start date
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim releaseDate As Date = CDate("01/01/0001")

        If Date.Compare(oLustEvent.Confirmed, CDate("01/01/0001")) = 0 Then
            releaseDate = oLustEvent.Started
        Else
            releaseDate = oLustEvent.Confirmed
        End If
        If oFeeInvoice.GetOwnerFeeBalanceOn(releaseDate, , oLustEvent.FacilityID) > 0 AndAlso (oLustEvent.Confirmed.Equals(DBNull.Value) OrElse oFeeInvoice.DueDate <= oLustEvent.Confirmed) Then
            GetTFAnswer4 = "N"
        Else
            GetTFAnswer4 = "Y"
        End If
        'If oFeeInvoice.GetCurrentBalance_Facility(oLustEvent.FacilityID) > 0 Then
        '    GetTFAnswer4 = "N"
        'Else
        '    GetTFAnswer4 = "Y"
        'End If
    End Function

    Private Function GetTFAnswer7() As String
        'Query the Lust Module Tank No. That Leaked and using that number, query the 
        'Notification Module Substance for that tank number.  If the tank substance is 
        'hazardous substance, used oil, unknown, other, not listed – mark “no”  otherwise, 
        'mark “yes”.  

        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim strResponse As String = "N"
        Try
            For Each ugrow In ugTankandPipes.Rows
                If ugrow.Band.Index = 0 Then
                    If ugrow.Cells("Status").Text.IndexOf("Currently In Use") > -1 Then
                        If Not ugrow.Cells("Substance").Value Is DBNull.Value Then
                            If Not (ugrow.Cells("Substance").Text.IndexOf("Hazardous Substance") > -1 Or _
                               ugrow.Cells("Substance").Text.IndexOf("Used Oil") > -1 Or _
                               ugrow.Cells("Substance").Text.IndexOf("Unknown") > -1 Or _
                               ugrow.Cells("Substance").Text.IndexOf("Other") > -1 Or _
                               ugrow.Cells("Substance").Text.IndexOf("Not Listed") > -1) Then
                                strResponse = "Y"
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error with TF Trustfund Question 7:  " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            GetTFAnswer7 = strResponse
        End Try
    End Function


    Private Function GetTFAnswer9() As String
        'At the time of the release, was the site in compliance with the 
        'Compliance and Enforcement Division? 
        Dim inspCitation As New MUSTER.BusinessLogic.pInspectionCitation
        Dim strResponse As String = "Y"

        Try
            If inspCitation.CheckCitationExists(oLustEvent.ReportDate, oLustEvent.FacilityID, False, , True, False, False) Then
                strResponse = "N"
            Else
                strResponse = "Y"
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error with TF Trustfund Question 9:  " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            GetTFAnswer9 = strResponse
        End Try
        'GetTFAnswer9 = "X"
    End Function


    Private Function GetTFAnswer11() As String
        'Do the tanks and lines have corrosion protection?
        ' logic: substandard in tos/tosi rules from registration
        'Dim bolCPMet As Boolean = False
        Dim bolCitationExists As Boolean = False
        Dim nTotalTankCount, nTotalPipeCount, nTankCPNotReq, nPipeCPNotReq, nTankCPReq, nPipeCPReq As Integer
        Dim nTankCPMet, nTankCPNotMet, nPipeCPMet, nPipeCPNotMet As Integer
        Dim dtPre88 As Date = CDate("12/22/1988")
        Dim dtPre98 As Date = CDate("12/22/1998")

        Try
            Dim inspCitation As New MUSTER.BusinessLogic.pInspectionCitation
            ' 10 - citaiton id for corrosion protection not maintained
            bolCitationExists = inspCitation.CheckCitationExists(CDate("01/01/0001"), oLustEvent.FacilityID, False, 10)

            nTotalTankCount = 0
            nTotalPipeCount = 0
            nTankCPNotReq = 0
            nPipeCPNotReq = 0
            nTankCPReq = 0
            nPipeCPReq = 0
            nTankCPMet = 0
            nTankCPNotMet = 0
            nPipeCPMet = 0
            nPipeCPNotMet = 0

            ' no need to check for null for the foll fields - using isnull(field,0) in db view
            ' TANKMATDESC, TANKMODDESC, PIPE_MAT_DESC, PIPE_MOD_DESC, PIPE_CP_TYPE

            ' bug #2160
            ' rules changed - third time
            ' ciu - tank/pipe
            ' installed before 12-22-1988 cp not req
            ' installed after 12-22-1988 cp req
            ' 
            ' tosi/pou - tank/pipe
            ' installed before 12-22-1988, dlu before 12-22-1988 cp not required
            ' installed before 12-22-1988, dlu before 12-22-1998 cp not required
            ' installed before 12-22-1988, dlu after 12-22-1998 cp required
            ' installed after 12-22-1988, dlu before 12-22-1998 cp required
            ' installed after 12-22-1988, dlu after 12-22-1998 cp required
            For Each ugrowTank As Infragistics.Win.UltraWinGrid.UltraGridRow In ugTankandPipes.Rows
                nTotalTankCount += 1
                'If ugrowTank.Cells("Included").Value = True Then

                ' New rules
                If ugrowTank.Cells("Status").Text.IndexOf("Currently In Use") > -1 Or _
                    ugrowTank.Cells("Status").Text.IndexOf("Temporarily Out of Service Indefinitely") > -1 Or _
                    ugrowTank.Cells("Status").Text.IndexOf("Permanently Out of Use") > -1 Then

                    If ugrowTank.Cells("Status").Text.IndexOf("Permanently Out of Use") = -1 Then
                        If Not ugrowTank.Cells("INSTALLED").Value Is DBNull.Value Then
                            nTankCPReq += 1
                            If ((ugrowTank.Cells("TANKMATDESC").Value = 344 Or ugrowTank.Cells("TANKMATDESC").Value = 350 Or ugrowTank.Cells("TANKMATDESC").Value = 351) And _
                                (ugrowTank.Cells("TANKMODDESC").Value = 414 Or ugrowTank.Cells("TANKMODDESC").Value = 0)) Or _
                                bolCitationExists Then
                                nTankCPNotMet += 1
                            Else
                                nTankCPMet += 1
                            End If ' if tank mat desc = ...
                        End If
                    Else
                        If Not ugrowTank.Cells("INSTALLED").Value Is DBNull.Value Then
                            If Date.Compare(ugrowTank.Cells("INSTALLED").Value, dtPre88) < 0 Then
                                If Not ugrowTank.Cells("DATELASTUSED").Value Is DBNull.Value Then
                                    If Date.Compare(ugrowTank.Cells("DATELASTUSED").Value, dtPre88) < 0 Then
                                        nTankCPNotReq += 1
                                    Else
                                        If Date.Compare(ugrowTank.Cells("DATELASTUSED").Value, dtPre98) < 0 Then
                                            nTankCPNotReq += 1
                                        Else
                                            nTankCPReq += 1
                                        End If
                                    End If
                                Else
                                    ' cp req irrespective of date last used pre / post 98
                                    nTankCPReq += 1
                                End If
                            Else
                                ' cp req irrespective of date last used pre / post 98
                                nTankCPReq += 1
                            End If
                            If ((ugrowTank.Cells("TANKMATDESC").Value = 344 Or ugrowTank.Cells("TANKMATDESC").Value = 350 Or ugrowTank.Cells("TANKMATDESC").Value = 351) And _
                                (ugrowTank.Cells("TANKMODDESC").Value = 414 Or ugrowTank.Cells("TANKMODDESC").Value = 0)) Or _
                                bolCitationExists Then
                                nTankCPNotMet += 1
                            Else
                                nTankCPMet += 1
                            End If ' if tank mat desc = ...
                        End If
                    End If

                End If 'if ciu/tosi/pou

                'If ugrowTank.Cells("Status").Text.IndexOf("Currently In Use") > -1 Or _
                '    ugrowTank.Cells("Status").Text.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then
                '    If ((ugrowTank.Cells("TANKMATDESC").Value = 344 Or ugrowTank.Cells("TANKMATDESC").Value = 350 Or ugrowTank.Cells("TANKMATDESC").Value = 351) And _
                '        (ugrowTank.Cells("TANKMODDESC").Value = 414 Or ugrowTank.Cells("TANKMODDESC").Value = 0)) Or _
                '        bolCitationExists Then
                '        bolCPMet = False
                '    Else
                '        bolCPMet = True
                '    End If ' if tank mat desc = ...
                'ElseIf ugrowTank.Cells("Status").Text.IndexOf("Permanently Out of Use") > -1 Then
                '    If Not ugrowTank.Cells("DATELASTUSED").Value Is DBNull.Value Then
                '        If Date.Compare(ugrowTank.Cells("DATELASTUSED").Value, dtPre88) >= 0 Then
                '            nPOUPre88TankCount += 1
                '        End If
                '    End If
                'End If ' if tank status = ciu/tosi
                'End If ' if included

                'If bolCPMet Then Exit For

                If Not ugrowTank.ChildBands Is Nothing Then
                    For Each ugrowPipe As Infragistics.Win.UltraWinGrid.UltraGridRow In ugrowTank.ChildBands(0).Rows
                        nTotalPipeCount += 1

                        ' New rules
                        If ugrowPipe.Cells("Status").Text.IndexOf("Currently In Use") > -1 Or _
                            ugrowPipe.Cells("Status").Text.IndexOf("Temporarily Out of Service Indefinitely") > -1 Or _
                            ugrowPipe.Cells("Status").Text.IndexOf("Permanently Out of Use") > -1 Then
                            If ugrowPipe.Cells("Status").Text.IndexOf("Permanently Out of Use") = -1 Then
                                If Not ugrowPipe.Cells("INSTALLED").Value Is DBNull.Value Then
                                    nPipeCPReq += 1
                                    If ((ugrowPipe.Cells("PIPE_MAT_DESC").Value = 250 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 253 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 251 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 256) And _
                                        (ugrowPipe.Cells("PIPE_MOD_DESC").Value <> 261 And ugrowPipe.Cells("PIPE_MOD_DESC").Value <> 260)) Or _
                                        (bolCitationExists And ugrowPipe.Cells("PIPE_CP_TYPE").Value = 478) Then
                                        nPipeCPNotMet += 1
                                    Else
                                        nPipeCPMet += 1
                                    End If ' if PIPE_MAT_DESC = ...
                                End If
                            Else
                                If Not ugrowPipe.Cells("INSTALLED").Value Is DBNull.Value Then
                                    If Date.Compare(ugrowPipe.Cells("INSTALLED").Value, dtPre88) < 0 Then
                                        If Not ugrowPipe.Cells("DATELASTUSED").Value Is DBNull.Value Then
                                            If Date.Compare(ugrowPipe.Cells("DATELASTUSED").Value, dtPre88) < 0 Then
                                                nPipeCPNotReq += 1
                                            Else
                                                If Date.Compare(ugrowPipe.Cells("DATELASTUSED").Value, dtPre98) < 0 Then
                                                    nPipeCPNotReq += 1
                                                Else
                                                    nPipeCPReq += 1
                                                End If
                                            End If
                                        Else
                                            ' cp req irrespective of date last used pre / post 98
                                            nPipeCPReq += 1
                                        End If
                                    Else
                                        ' cp req irrespective of date last used pre / post 98
                                        nPipeCPReq += 1
                                    End If
                                    If ((ugrowPipe.Cells("PIPE_MAT_DESC").Value = 250 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 253 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 251 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 256) And _
                                        (ugrowPipe.Cells("PIPE_MOD_DESC").Value <> 261 And ugrowPipe.Cells("PIPE_MOD_DESC").Value <> 260)) Or _
                                        (bolCitationExists And ugrowPipe.Cells("PIPE_CP_TYPE").Value = 478) Then
                                        nPipeCPNotMet += 1
                                    Else
                                        nPipeCPMet += 1
                                    End If ' if PIPE_MAT_DESC = ...
                                End If
                            End If
                        End If ' if ciu/tosi/pou

                        'If ugrowPipe.Cells("Included").Value = True Then
                        'If ugrowPipe.Cells("Status").Text.IndexOf("Currently In Use") > -1 Or _
                        '    ugrowPipe.Cells("Status").Text.IndexOf("Temporarily Out of Service Indefinitely") > -1 Then
                        '    If ((ugrowPipe.Cells("PIPE_MAT_DESC").Value = 250 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 253 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 251 Or ugrowPipe.Cells("PIPE_MAT_DESC").Value = 256) And _
                        '        (ugrowPipe.Cells("PIPE_MOD_DESC").Value <> 261 Or ugrowPipe.Cells("PIPE_MOD_DESC").Value <> 260)) Or _
                        '        (bolCitationExists And ugrowPipe.Cells("PIPE_CP_TYPE").Value = 478) Then
                        '        bolCPMet = False
                        '    Else
                        '        bolCPMet = True
                        '    End If ' if PIPE_MAT_DESC = ...
                        'ElseIf ugrowPipe.Cells("Status").Text.IndexOf("Permanently Out of Use") > -1 Then
                        '    If Not ugrowPipe.Cells("DATELASTUSED").Value Is DBNull.Value Then
                        '        If Date.Compare(ugrowPipe.Cells("DATELASTUSED").Value, dtPre88) >= 0 Then
                        '            nPOUPre88PipeCount += 1
                        '        End If
                        '    End If
                        'End If ' if pipe status = ciu/tosi
                        'End If ' if included

                        'If bolCPMet Then Exit For

                    Next ' pipe
                End If ' if not childband is null

                'If bolCPMet Then Exit For

            Next ' tank
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error with TF Trustfund Question 11:  " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            If nTotalTankCount = nTankCPNotReq And nTotalPipeCount = nPipeCPNotReq Then
                GetTFAnswer11 = "A"
            Else
                If nTankCPMet >= nTankCPReq And nPipeCPMet >= nPipeCPReq Then
                    GetTFAnswer11 = "Y"
                Else
                    GetTFAnswer11 = "N"
                End If
            End If
            'GetTFAnswer11 = IIf(bolCPMet, "Y", _
            '                    IIf(nTotalTankCount = nPOUPre88TankCount And _
            '                        nTotalPipeCount = nPOUPre88PipeCount And _
            '                        nTotalTankCount <> 0 And nTotalPipeCount <> 0, "A", "N"))
        End Try
        'GetTFAnswer11 = "X"
    End Function


    Private Function GetTFAnswer12() As String
        'Does the UST system have required spill prevention?
        Dim ugrowTank As Infragistics.Win.UltraWinGrid.UltraGridRow
        'Dim oTankInfo As MUSTER.Info.TankInfo
        'Dim strResponse As String

        Dim nTotalTankCount, nTankCPNotReq, nTankCPReq As Integer
        Dim nTankCPMet, nTankCPNotMet As Integer
        Dim dtPre88 As Date = CDate("12/22/1988")
        Dim dtPre98 As Date = CDate("12/22/1998")

        'Dim nTotalTankCount, nPOUPre88TankCount As Integer
        'Dim dtPre88 As Date = CDate("12/22/88")

        'strResponse = "A"
        Try
            For Each ugrowTank In ugTankandPipes.Rows
                nTotalTankCount += 1

                ' New rules
                If ugrowTank.Cells("Status").Text.IndexOf("Currently In Use") > -1 Or _
                    ugrowTank.Cells("Status").Text.IndexOf("Temporarily Out of Service Indefinitely") > -1 Or _
                    ugrowTank.Cells("Status").Text.IndexOf("Permanently Out of Use") > -1 Then

                    If ugrowTank.Cells("Status").Text.IndexOf("Permanently Out of Use") = -1 Then
                        If Not ugrowTank.Cells("INSTALLED").Value Is DBNull.Value Then
                            nTankCPReq += 1
                            If Not ugrowTank.Cells("SPILLINSTALLED").Value Is DBNull.Value Then
                                If ugrowTank.Cells("SPILLINSTALLED").Value Then
                                    nTankCPMet += 1
                                Else
                                    nTankCPNotMet += 1
                                End If
                            Else
                                nTankCPNotMet += 1
                            End If
                        End If
                    Else
                        If Not ugrowTank.Cells("INSTALLED").Value Is DBNull.Value Then
                            If Date.Compare(ugrowTank.Cells("INSTALLED").Value, dtPre88) < 0 Then
                                If Not ugrowTank.Cells("DATELASTUSED").Value Is DBNull.Value Then
                                    If Date.Compare(ugrowTank.Cells("DATELASTUSED").Value, dtPre88) < 0 Then
                                        nTankCPNotReq += 1
                                    Else
                                        If Date.Compare(ugrowTank.Cells("DATELASTUSED").Value, dtPre98) < 0 Then
                                            nTankCPNotReq += 1
                                        Else
                                            nTankCPReq += 1
                                        End If
                                    End If
                                Else
                                    ' cp req irrespective of date last used pre / post 98
                                    nTankCPReq += 1
                                End If
                            Else
                                ' cp req irrespective of date last used pre / post 98
                                nTankCPReq += 1
                            End If
                            If Not ugrowTank.Cells("SPILLINSTALLED").Value Is DBNull.Value Then
                                If ugrowTank.Cells("SPILLINSTALLED").Value Then
                                    nTankCPMet += 1
                                Else
                                    nTankCPNotMet += 1
                                End If
                            Else
                                nTankCPNotMet += 1
                            End If
                        End If
                    End If

                End If 'if ciu/tosi/pou

                'If ugrow.Band.Index = 0 Then
                'If ugrow.Cells("Included").Value = True Then
                'If ugrow.Cells("Status").Text.IndexOf("Currently In Use") > -1 Or _
                '    ugrow.Cells("Status").Text.IndexOf("Temporarily Out of Service") > -1 Then ' CIU / TOS / TOSI
                '    If Not ugrow.Cells("SPILLINSTALLED").Value Is DBNull.Value Then
                '        If ugrow.Cells("SPILLINSTALLED").Value Then
                '            strResponse = "Y"
                '            Exit For
                '        Else
                '            strResponse = "N"
                '            Exit For
                '        End If
                '    Else
                '        strResponse = "N"
                '        Exit For
                '    End If
                'ElseIf ugrow.Cells("Status").Text.IndexOf("Permanently Out of Use") > -1 Then
                '    If Not ugrow.Cells("DATELASTUSED").Value Is DBNull.Value Then
                '        If Date.Compare(ugrow.Cells("DATELASTUSED").Value, dtPre88) >= 0 Then
                '            nPOUPre88TankCount += 1
                '        End If
                '    End If
                'End If
                'For Each oTankInfo In oOwner.Facility.TankCollection.Values
                '    If CInt(ugrow.Cells("TANK ID").Value) = oTankInfo.TankId Then
                '        If oTankInfo.SpillInstalled Then
                '            strResponse = "Y"
                '        End If
                '    End If
                'Next
                'End If
                'End If
            Next
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error with TF Trustfund Question 12:  " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally

            If nTotalTankCount = nTankCPNotReq Then


                'If nTankCPMet = nTotalTankCount Then
                'GetTFAnswer12 = "Y"
                'Else
                'GetTFAnswer12 = "Y"
                'End If
                GetTFAnswer12 = "A"
            Else
            If nTankCPMet >= nTankCPReq Then
                GetTFAnswer12 = "Y"
            Else
                GetTFAnswer12 = "N"
            End If
            End If
            'GetTFAnswer12 = IIf(strResponse = "A", _
            '                    IIf(nTotalTankCount = nPOUPre88TankCount And _
            '                        nTotalTankCount <> 0, "A", "N"), strResponse)
        End Try

    End Function
    Private Function GetTFAnswer13() As String
        'Does the UST system have required overfill prevention?
        Dim ugrowTank As Infragistics.Win.UltraWinGrid.UltraGridRow
        'Dim oTankInfo As MUSTER.Info.TankInfo
        'Dim strResponse As String
        'Dim nTotalTankCount, nPOUPre88TankCount As Integer
        'Dim dtPre88 As Date = CDate("12/22/88")
        'strResponse = "A"

        Dim nTotalTankCount, nTankCPNotReq, nTankCPReq As Integer
        Dim nTankCPMet, nTankCPNotMet As Integer
        Dim dtPre88 As Date = CDate("12/22/1988")
        Dim dtPre98 As Date = CDate("12/22/1998")

        Try
            For Each ugrowTank In ugTankandPipes.Rows
                nTotalTankCount += 1

                ' New rules
                If ugrowTank.Cells("Status").Text.IndexOf("Currently In Use") > -1 Or _
                    ugrowTank.Cells("Status").Text.IndexOf("Temporarily Out of Service Indefinitely") > -1 Or _
                    ugrowTank.Cells("Status").Text.IndexOf("Permanently Out of Use") > -1 Then

                    If ugrowTank.Cells("Status").Text.IndexOf("Permanently Out of Use") = -1 Then
                        If Not ugrowTank.Cells("INSTALLED").Value Is DBNull.Value Then
                            nTankCPReq += 1

                            If Not ugrowTank.Cells("SPILLINSTALLED").Value Is DBNull.Value Then
                                If ugrowTank.Cells("SPILLINSTALLED").Value Then
                                    nTankCPMet += 1
                                Else
                                    nTankCPNotMet += 1
                                End If
                            Else
                                nTankCPNotMet += 1
                            End If
                        End If
                    Else

                        If Not ugrowTank.Cells("INSTALLED").Value Is DBNull.Value Then
                            If Date.Compare(ugrowTank.Cells("INSTALLED").Value, dtPre88) < 0 Then
                                If Not ugrowTank.Cells("DATELASTUSED").Value Is DBNull.Value Then
                                    If Date.Compare(ugrowTank.Cells("DATELASTUSED").Value, dtPre88) < 0 Then
                                        nTankCPNotReq += 1
                                    Else
                                        If Date.Compare(ugrowTank.Cells("DATELASTUSED").Value, dtPre98) < 0 Then
                                            nTankCPNotReq += 1
                                        Else
                                            nTankCPReq += 1
                                        End If
                                    End If
                                Else
                                    ' cp req irrespective of date last used pre / post 98
                                    nTankCPReq += 1
                                End If
                            Else
                                ' cp req irrespective of date last used pre / post 98
                                nTankCPReq += 1
                            End If
                            If Not ugrowTank.Cells("SPILLINSTALLED").Value Is DBNull.Value Then
                                If ugrowTank.Cells("SPILLINSTALLED").Value Then
                                    nTankCPMet += 1
                                Else
                                    nTankCPNotMet += 1
                                End If
                            Else
                                nTankCPNotMet += 1
                            End If
                        End If
                    End If

                End If 'if ciu/tosi/pou

                'If ugrow.Band.Index = 0 Then
                'If ugrow.Cells("Included").Value = True Then
                'If ugrow.Cells("Status").Text.IndexOf("Currently In Use") > -1 Or _
                '    ugrow.Cells("Status").Text.IndexOf("Temporarily Out of Service") > -1 Then ' CIU / TOS / TOSI
                '    If Not ugrow.Cells("OVERFILLINSTALLED").Value Is DBNull.Value Then
                '        If ugrow.Cells("OVERFILLINSTALLED").Value Then
                '            strResponse = "Y"
                '            Exit For
                '        Else
                '            strResponse = "N"
                '            Exit For
                '        End If
                '    Else
                '        strResponse = "N"
                '        Exit For
                '    End If
                'ElseIf ugrow.Cells("Status").Text.IndexOf("Permanently Out of Use") > -1 Then
                '    If Not ugrow.Cells("DATELASTUSED").Value Is DBNull.Value Then
                '        If Date.Compare(ugrow.Cells("DATELASTUSED").Value, dtPre88) >= 0 Then
                '            nPOUPre88TankCount += 1
                '        End If
                '    End If
                'End If
                'For Each oTankInfo In oOwner.Facility.TankCollection.Values
                '    If CInt(ugrow.Cells("TANK ID").Value) = oTankInfo.TankId Then
                '        If oTankInfo.OverFillInstalled Then
                '            strResponse = "Y"
                '        End If
                '    End If
                'Next
                'End If
                'End If
            Next

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error with TF Trustfund Question 13:  " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            If nTotalTankCount = nTankCPNotReq Then
                'If nTankCPMet = nTotalTankCount Then
                'GetTFAnswer13 = "Y"
                'Else
                'GetTFAnswer13 = "Y"
                'End If
                GetTFAnswer13 = "A"
            Else
                If nTankCPMet >= nTankCPReq Then
                    GetTFAnswer13 = "Y"
                Else
                    GetTFAnswer13 = "N"
                End If
            End If
            'GetTFAnswer13 = IIf(strResponse = "A", _
            '                    IIf(nTotalTankCount = nPOUPre88TankCount And _
            '                        nTotalTankCount <> 0, "A", "N"), strResponse)
        End Try


    End Function

#End Region

#Region " Comments "
    Private Sub btnViewModifyComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewModifyComment.Click
        Try
            If oLustEvent.ID <> 0 Then
                CommentsMaintenance(sender, e)
            Else
                MsgBox("You Must Save The Lust Event Prior To Entering Comments.")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub txtEligibilityComments_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEligibilityComments.TextChanged
        If bolLoading Then Exit Sub
        oLustEvent.ELIGIBITY_COMMENTS = txtEligibilityComments.Text

    End Sub
#End Region

  
#Region "Contacts"
#Region "Button and Change Events"
    Dim strFilterString As String = String.Empty
    Private Sub ugOwnerContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugOwnerContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ugOwnerContacts)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub tbCtrlOwner_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlOwner.Click
        If bolLoading Then Exit Sub
        Try
            Select Case tbCtrlOwner.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If Me.lblOwnerIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, Integer.Parse(Me.lblOwnerIDValue.Text), UIUtilsGen.EntityTypes.Owner)
                    End If
                Case tbPageOwnerDocuments.Name
                    UCOwnerDocuments.LoadDocumentsGrid(Integer.Parse(Me.lblOwnerIDValue.Text), UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Technical)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerAddSearchContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAddSearchContact.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                objCntSearch = New ContactSearch(oOwner.ID, 9, "Technical", pConStruct)
            Else
                objCntSearch = New ContactSearch(oOwner.Facility.ID, 6, "Technical", pConStruct)
            End If
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
            'objCntSearch.BringToFront()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerModifyContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerModifyContact.Click
        Try
            ModifyContact(ugOwnerContacts)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub btnOwnerAssociateContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAssociateContact.Click
        Try
            If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                AssociateContact(ugOwnerContacts, oOwner.ID, 9)
            Else
                AssociateContact(ugOwnerContacts, oOwner.Facility.ID, 6)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub btnOwnerDeleteContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerDeleteContact.Click
        Try
            If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                DeleteContact(ugOwnerContacts, oOwner.ID)
            Else
                DeleteContact(ugOwnerContacts, oOwner.Facility.ID)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerContactClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.Close()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub chkOwnerShowActiveOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowActiveOnly.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkOwnerShowContactsforAllModules_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowContactsforAllModules.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkOwnerShowRelatedContacts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOwnerShowRelatedContacts.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub SetLustFilter()
        Dim strMode As String = String.Empty
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strEntities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityAssocIDs As String = String.Empty
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Try
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' strFilterString = String.Empty
            If chkERACShowActive.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If

            If chkERACShowContactsforAllModules.Checked Then
                ' User has the ability to view the contacts associated for the entity in other modules
                Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(oOwner.Facility.ID.ToString)
                nEntityID = oOwner.Facility.ID

                nEntityType = 6
                strEntityAssocIDs = String.Format("{0}FOR:{1}", strFilterForAllModules, oLustEvent.ID.ToString)
                nModuleID = 0
            Else
                nEntityID = oLustEvent.ID
                nEntityType = 7
                nModuleID = 614
            End If

            If chkERACShowRelatedContacts.Checked Then
                strEntities = strLustEventIdTags
                nRelatedEntityType = 7
            End If

            UIUtilsGen.LoadContacts(Me.ugERACContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, strEntityAssocIDs, nRelatedEntityType)



            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'strFilterString = String.Empty
            'If chkERACShowActive.Checked Then
            '    strFilterString = "(ACTIVE = 1"
            'Else
            '    strFilterString = "("
            'End If

            'If chkERACShowContactsforAllModules.Checked Then
            '    ' User has the ability to view the contacts associated for the entity in other modules
            '    'User has the ability to view the contacts associated for the entity in other modules
            '    Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(oOwner.Facility.ID.ToString)
            '    If strFilterString = "(" Then
            '        strFilterString += "ENTITYID = " + oOwner.Facility.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '    Else
            '        strFilterString += "AND ENTITYID = " + oOwner.Facility.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '    End If
            '    'Else
            '    '    If strFilterString = "(" Then
            '    '        strFilterString += "ENTITYID = " + strEntityID
            '    '    Else
            '    '        strFilterString += "AND ENTITYID = " + strEntityID
            '    '    End If
            '    'End If
            '    'If strFilterString = "(" Then
            '    '    strFilterString += "ENTITYID = " + oOwner.Facility.ID.ToString + " OR ENTITYID = " + oLustEvent.ID.ToString
            '    'Else
            '    '    strFilterString += "AND ENTITYID = " + oOwner.Facility.ID.ToString + " OR ENTITYID = " + oLustEvent.ID.ToString
            '    'End If
            'Else
            '    If strFilterString = "(" Then
            '        strFilterString += " MODULEID = 614 And ENTITYID = " + oLustEvent.ID.ToString
            '    Else
            '        strFilterString += " AND MODULEID = 614 And ENTITYID = " + oLustEvent.ID.ToString
            '    End If
            'End If

            'If chkERACShowRelatedContacts.Checked Then
            '    strFilterString += " OR " + IIf(Not strLustEventIdTags = String.Empty, " ENTITYID in (" + strLustEventIdTags + "))", "")
            'Else
            '    strFilterString += ")"
            'End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugERACContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub SetFilter()
        Dim strMode As String = String.Empty
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strEntities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityAssocIDs As String = String.Empty
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                nEntityID = oOwner.ID
                nEntityType = 9
                strMode = "OWNER"
            ElseIf tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEFACILITYDETAIL" Then
                nEntityID = oOwner.Facility.ID
                nEntityType = 6
            Else
                nEntityID = oLustEvent.ID
            End If

            If chkOwnerShowActiveOnly.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If

            If chkOwnerShowContactsforAllModules.Checked Then
                'User has the ability to view the contacts associated for the entity in other modules
                If strMode <> "OWNER" Then
                    Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(oOwner.Facility.ID.ToString)
                    strEntityAssocIDs = strFilterForAllModules
                    nModuleID = 0
                End If
            Else
                nModuleID = 614
            End If
            If chkOwnerShowRelatedContacts.Checked Then
                strEntities = strFacilityIdTags
                nRelatedEntityType = 6
            End If

            UIUtilsGen.LoadContacts(ugOwnerContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, strEntityAssocIDs, nRelatedEntityType)



            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'strFilterString = String.Empty
            'Dim strEntityID As String
            'If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
            '    strEntityID = oOwner.ID.ToString
            '    strMode = "OWNER"
            'ElseIf tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEFACILITYDETAIL" Then
            '    strEntityID = oOwner.Facility.ID.ToString
            'Else
            '    strEntityID = oLustEvent.ID.ToString
            'End If

            'If chkOwnerShowActiveOnly.Checked Then
            '    strFilterString = "(ACTIVE = 1"
            'Else
            '    strFilterString = "("
            'End If

            'If chkOwnerShowContactsforAllModules.Checked Then

            '    'User has the ability to view the contacts associated for the entity in other modules
            '    If strMode <> "OWNER" Then
            '        Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(oOwner.Facility.ID.ToString)
            '        If strFilterString = "(" Then
            '            strFilterString += "ENTITYID = " + strEntityID + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '        Else
            '            strFilterString += "AND ENTITYID = " + strEntityID + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '        End If
            '    Else
            '        If strFilterString = "(" Then
            '            strFilterString += "ENTITYID = " + strEntityID
            '        Else
            '            strFilterString += "AND ENTITYID = " + strEntityID
            '        End If
            '    End If

            'Else
            '    If strFilterString = "(" Then
            '        strFilterString += " MODULEID = 614 And ENTITYID = " + strEntityID
            '    Else
            '        strFilterString += " AND MODULEID = 614 And ENTITYID = " + strEntityID
            '    End If
            'End If
            'If chkOwnerShowRelatedContacts.Checked Then
            '    strFilterString += " OR " + IIf(Not strFacilityIdTags = String.Empty, " ENTITYID in (" + strFacilityIdTags + "))", "")
            'Else
            '    strFilterString += ")"
            'End If
            ''If chkOwnerShowRelatedContacts.Checked Then
            ''    If strFilterString = String.Empty Then
            ''        strFilterString = " (ENTITYID = " + strEntityID + IIf(Not strFacilityIdTags = String.Empty, " OR ENTITYID in (" + strFacilityIdTags + ")", "") + ")"
            ''    Else
            ''        strFilterString += " AND (ENTITYID = " + strEntityID + IIf(Not strFacilityIdTags = String.Empty, " OR ENTITYID in (" + strFacilityIdTags + ")", "") + ")"
            ''    End If
            ''Else
            ''    strFilterString += ""
            ''End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugOwnerContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "LUST Contacts"
    Private Sub chkERACShowContactsforAllModules_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkERACShowContactsforAllModules.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetLustFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkERACShowRelatedContacts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkERACShowRelatedContacts.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetLustFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkERACShowActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkERACShowActive.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetLustFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnERACContactAddorSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnERACContactAddorSearch.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            objCntSearch = New ContactSearch(oLustEvent.ID, 7, "Technical", pConStruct)
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub btnERACContactModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnERACContactModify.Click
        Try
            ModifyContact(ugERACContacts)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnERACContactDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnERACContactDelete.Click
        Try
            DeleteContact(ugERACContacts, oLustEvent.ID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnERACContactAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnERACContactAssociate.Click
        Try
            AssociateContact(ugERACContacts, oLustEvent.ID, 7)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub btnERACContactClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub ERACSetFilter()
        Try
            strFilterString = String.Empty
            Dim strEntityID As String

            strEntityID = oLustEvent.ID.ToString

            If chkERACShowActive.Checked Then
                strFilterString = "ACTIVE = 1"
            Else
                strFilterString = ""
            End If

            If chkERACShowContactsforAllModules.Checked Then
                strFilterString += ""
            Else
                If strFilterString = String.Empty Then
                    strFilterString = " MODULEID = 614 And ENTITYID = " + strEntityID
                Else
                    strFilterString += " AND MODULEID = 614 And ENTITYID = " + strEntityID
                End If
            End If

            If Not chkERACShowRelatedContacts.Checked Then
                strFilterString += ""
            End If

            dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            ugERACContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Common Functions"

    Private Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer)

        Try

            UIUtilsGen.LoadContacts(ugGrid, EntityID, EntityType, pConStruct, 614)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Function ModifyContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try

            UIUtilsGen.ModifyContact(ugGrid, 614, pConStruct)

            Me.Contact_ContactAdded()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer)
        Try
            UIUtilsGen.AssociateContact(ugGrid, nEntityID, nEntityType, 614, pConStruct)

            Me.Contact_ContactAdded()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer)
        Try

            UIUtilsGen.DeleteContact(ugGrid, nEntityID, 614, pConStruct)

            Me.Contact_ContactAdded()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

#End Region
#Region "Close Events"
    Private Sub Search_ContactAdded() Handles objCntSearch.ContactAdded
        If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGELUSTEVENT" Then
            LoadContacts(ugERACContacts, oLustEvent.ID, 7)
            chkERACShowContactsforAllModules.Checked = False
        Else
            If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                LoadContacts(ugOwnerContacts, oOwner.ID, 9)
            Else
                LoadContacts(ugOwnerContacts, oOwner.Facility.ID, 6)
            End If
            chkOwnerShowContactsforAllModules.Checked = False
        End If
    End Sub
    Private Sub Contact_ContactAdded()
        If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGELUSTEVENT" Then
            LoadContacts(ugERACContacts, oLustEvent.ID, 7)
            chkERACShowContactsforAllModules.Checked = False
        Else
            If tbCntrlTechnical.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                LoadContacts(ugOwnerContacts, oOwner.ID, 9)
            Else
                LoadContacts(ugOwnerContacts, oOwner.Facility.ID, 6)
            End If
            chkOwnerShowContactsforAllModules.Checked = False
        End If

    End Sub

    Private Sub objCntSearch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objCntSearch.Closing
        If Not objCntSearch Is Nothing Then
            objCntSearch = Nothing
        End If
    End Sub
#End Region

#End Region

#Region "Owner/Facility Comments"
    Private Sub CommentsMaintenance(Optional ByVal sender As System.Object = Nothing, Optional ByVal e As System.EventArgs = Nothing, Optional ByVal bolSetCounts As Boolean = False, Optional ByVal resetBtnColor As Boolean = False)
        Dim SC As ShowComments
        Dim nEntityID As Integer = 0
        Dim nEntityType As Integer = 0
        Dim strEntityName As String = String.Empty
        Dim oComments As MUSTER.BusinessLogic.pComments
        Dim bolEnableShowAllModules As Boolean = True
        Dim nCommentsCount As Integer = 0
        Try
            Select Case Me.tbCntrlTechnical.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    strEntityName = "Owner : " + CStr(oOwner.ID) + " " + Me.txtOwnerName.Text
                    oComments = oOwner.Comments
                    nEntityID = oOwner.ID
                    nEntityType = UIUtilsGen.EntityTypes.Owner
                Case tbPageFacilityDetail.Name
                    strEntityName = "Facility : " + CStr(oOwner.Facilities.ID) + " " + oOwner.Facilities.Name
                    oComments = oOwner.Facilities.Comments
                    nEntityID = oOwner.Facilities.ID
                    nEntityType = UIUtilsGen.EntityTypes.Facility
                Case tbPageLUSTEvent.Name
                    strEntityName = "Facility : " + CStr(oOwner.Facilities.ID) + " " + oOwner.Facilities.Name + ", Lust Event : " + CStr(oLustEvent.EVENTSEQUENCE)
                    oComments = New MUSTER.BusinessLogic.pComments
                    nEntityID = oLustEvent.ID
                    nEntityType = UIUtilsGen.EntityTypes.LUST_Event
                    bolEnableShowAllModules = False
                Case Else
                    Exit Sub
            End Select
            If Not resetBtnColor Then
                SC = New ShowComments(nEntityID, nEntityType, IIf(bolSetCounts, "", "Technical"), strEntityName, oComments, Me.Text, , bolEnableShowAllModules)
                If bolSetCounts Then
                    nCommentsCount = SC.GetCounts()
                Else
                    SC.ShowDialog()
                    nCommentsCount = IIf(SC.nCommentsCount <= 0, SC.GetCounts(), SC.nCommentsCount)
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
            ElseIf nEntityType = UIUtilsGen.EntityTypes.LUST_Event Then
                If nCommentsCount > 0 Then
                    btnViewModifyComment.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnViewModifyComment.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Flags"
#Region "UI Support Routines"
    Private Sub SF_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles SF.FlagAdded
        ' New event declared to allow registration to trigger
        '   container to check flag status.  This is bogus, but
        '   the current design which permits something other than
        '   MusterContainer to create a registration object (i.e.
        '   OwnerSearchResults creates it, so MusterContainer
        '   doesn't know about it) precludes use of events to notify
        '   MusterContainer of an request.  Therefore, the event FlagsChanged
        '   cannot be fired here and then caught by the MusterContainer.
        '   
        Dim MyFrm As MusterContainer
        MyFrm = Me.MdiParent
        'oEntity.GetEntity("Owner")
        If Not MyFrm Is Nothing Then
            Select Case Me.tbCntrlTechnical.SelectedTab.Name
                Case tbPageLUSTEvent.Name
                    MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, [Module], ParentFormText, entityID, entityType)
                Case Else
                    MyFrm.FlagsChanged(entityID, entityType, [Module], ParentFormText)
            End Select
        End If
        'If entityType = UIUtilsGen.EntityTypes.LUST_Event Then
        '    entityID = oLustEvent.FacilityID
        '    entityType = UIUtilsGen.EntityTypes.Facility
        'End If
        'MyFrm.FlagsChanged(entityID, entityType, [Module], ParentFormText)
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        Dim mc As MusterContainer = Me.MdiParent
        mc.RefreshCalendarInfo()
        mc.LoadDueToMeCalendar()
        mc.LoadToDoCalendar()
    End Sub
    Private Sub FlagMaintenance(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Select Case Me.tbCntrlTechnical.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    SF = New ShowFlags(oOwner.ID, UIUtilsGen.EntityTypes.Owner, "Technical")
                Case tbPageFacilityDetail.Name
                    SF = New ShowFlags(oOwner.Facilities.ID, UIUtilsGen.EntityTypes.Facility, "Technical")
                Case tbPageLUSTEvent.Name
                    If oLustEvent.ID <= 0 And oLustEvent.ID >= -100 Then
                        MsgBox("You Must Save The Lust Event Prior To Entering Flags.")
                        Exit Sub
                    End If
                    'SF = New ShowFlags(oOwner.Facilities.ID, UIUtilsGen.EntityTypes.Facility, "Technical")
                    SF = New ShowFlags(oLustEvent.ID, UIUtilsGen.EntityTypes.LUST_Event, "Technical", , , oLustEvent.EVENTSEQUENCE.ToString)
                Case Else
                    Exit Sub
            End Select
            SF.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub btnOwnerFlag_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerFlag.Click
        FlagMaintenance(sender, e)
        'Dim MyFrm As MusterContainer
        'MyFrm = Me.MdiParent
        'oEntity.GetEntity("Owner")
        'MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, 0, 0, 0, "Technical", Me.Text)
    End Sub
    Private Sub btnFacFlags_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacFlags.Click
        FlagMaintenance(sender, e)
        'Dim MyFrm As MusterContainer
        'MyFrm = Me.MdiParent
        'oEntity.GetEntity("Facility")
        'MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, Me.lblFacilityIDValue.Text, 0, 0, "Technical", Me.Text)
    End Sub
    Private Sub btnFlagsLustEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFlagsLustEvent.Click
        FlagMaintenance(sender, e)
    End Sub
#End Region
#End Region

    Private Sub CheckForSingleEvent_ByOwner(ByVal OwnerID As Int64)
        Dim EventID As Int64

        EventID = oLustEvent.CheckForSingleOpenLustEvent(OwnerID, 0)

        If EventID = 0 Then Exit Sub
        oLustEvent.Retrieve(EventID)
        If oLustEvent.FacilityID = 0 Then Exit Sub

        PopFacility(oLustEvent.FacilityID)

        SetupModifyLustEventForm(True, False)

        Me.tbCntrlTechnical.SelectedTab = Me.tbPageLUSTEvent

        btnViewModifyComment.Select()
        btnViewModifyComment.Focus()
    End Sub

    Private Sub CheckForSingleEvent_ByFacility(ByVal FacilityID As Int64)
        Dim EventID As Int64

        EventID = oLustEvent.CheckForSingleOpenLustEvent(0, FacilityID)

        If EventID = 0 Then Exit Sub

        oLustEvent.Retrieve(EventID)

        SetupModifyLustEventForm(True, False)

        Me.tbCntrlTechnical.SelectedTab = Me.tbPageLUSTEvent

        btnViewModifyComment.Select()
        btnViewModifyComment.Focus()
    End Sub

    Private Sub lblERACSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblERACSearch.Click
        strFromEracIrac = "ERAC"
        oCompanySearch = New CompanySearch("Technical")
        oCompanySearch.ShowDialog()
    End Sub

    Private Sub lblIRACSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblIRACSearch.Click
        strFromEracIrac = "IRAC"
        oCompanySearch = New CompanySearch
        oCompanySearch.ShowDialog()
    End Sub
    Private Sub CompanyDetails(ByVal nCompanyID As Integer, ByVal strCompanyName As String) Handles oCompanySearch.CompanyDetails
        Try
            If strFromEracIrac = "ERAC" Then
                txtERAC.Text = strCompanyName
                oLustEvent.ERAC = nCompanyID
            Else
                txtIRAC.Text = strCompanyName
                oLustEvent.IRAC = nCompanyID
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub LicenseeCompanyDetails(ByVal Licensee_id As Integer, ByVal company_id As Integer, ByVal Licensee_name As String, ByVal company_name As String) Handles oCompanySearch.LicenseeCompanyDetails
        Try
            If strFromEracIrac = "ERAC" Then
                txtERAC.Text = company_name
                oLustEvent.ERAC = company_id
            Else
                txtIRAC.Text = company_name
                oLustEvent.IRAC = company_id
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub lblPMHistorytt_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPMHistorytt.DoubleClick
        Dim strPMHIstory As String
        strPMHIstory = oLustEvent.LustEventPMHistory
        If IsDBNull(strPMHIstory) Then
            strPMHIstory = String.Empty
        End If
        If lblEventID.Text = "Add Event" Or strPMHIstory = String.Empty Then
            Exit Sub
        Else
            MsgBox(strPMHIstory, MsgBoxStyle.Information, "PM History")
        End If


    End Sub

    Private Sub btnGoToFinancial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGoToFinancial.Click
        Dim FinancialForm As Financial
        Dim localGUID As String
        Dim ChildForm As Windows.Forms.Form
        Dim fGUID As System.Guid
        Dim MyFrm As MusterContainer



        MyFrm = Me.MdiParent
        localGUID = MusterContainer.AppSemaphores.GetGUIDForRef("OwnerID", oOwner.ID, "Financial")

        If Not localGUID = String.Empty Then
            For Each ChildForm In MyFrm.MdiChildren
                If ChildForm.GetType.Name = "Financial" Then
                    fGUID = CType(ChildForm, Financial).MyGuid
                    If fGUID.ToString = localGUID Then
                        ChildForm.Activate()
                        FinancialForm = CType(ChildForm, Financial)
                        If FinancialForm.lblOwnerIDValue.Text.Trim = oOwner.ID.ToString Or (FinancialForm.lblFacilityIDValue.Text.Trim = lblFacilityIDValue.Text) Then
                            FinancialForm.PopulateFacilityInfo(Integer.Parse(lblFacilityIDValue.Text))
                            FinancialForm.bolFromTechnical = True
                            FinancialForm.LoadFinancialData(oFinancialEvent.ID)
                            Exit For
                        End If
                    End If
                End If
            Next
        Else
            FinancialForm = New Financial(oOwner, oOwner.ID, lblFacilityIDValue.Text, oFinancialEvent.ID, True)
        End If

        FinancialForm.Tag = "Financial"
        FinancialForm.MdiParent = Me.MdiParent
        FinancialForm.Show()
        FinancialForm.BringToFront()

    End Sub

    Private Sub tbCntrlFacility_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlFacility.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            Select Case tbCntrlFacility.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If Me.lblFacilityIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility)
                    End If
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Technical)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub lnkEnsite_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkEnsite.LinkClicked
        Try
            lnkEnsite.LinkVisited = True
            If oOwner.Facilities.AIID > 0 Then
                'System.Diagnostics.Process.Start("http://opcweb/ensearch/agency_interest_details.aspx?ai=19252")
                System.Diagnostics.Process.Start("http://opcweb/ensearch/agency_interest_details.aspx?ai=" + oOwner.Facilities.AIID.ToString)
            Else
                MsgBox("No Facility AIID")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugERACContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugERACContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ugERACContacts)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#Region "Envelopes and Labels"
    Private Sub btnTecFacEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTecFacEnvelopes.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = oOwner.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = oOwner.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = oOwner.Facilities.FacilityAddresses.City
            arrAddress(3) = oOwner.Facilities.FacilityAddresses.State
            arrAddress(4) = oOwner.Facilities.FacilityAddresses.Zip
            'strAddress = oOwner.Facilities.FacilityAddresses.AddressLine1 + "," + oOwner.Facilities.FacilityAddresses.AddressLine2 + "," + oOwner.Facilities.FacilityAddresses.City + "," + oOwner.Facilities.FacilityAddresses.State + "," + oOwner.Facilities.FacilityAddresses.Zip
            If oOwner.Facilities.FacilityAddresses.AddressId > 0 And oOwner.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(Me.txtFacilityName.Text, arrAddress, "TEC", oOwner.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnTecFacLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTecFacLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = oOwner.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = oOwner.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = oOwner.Facilities.FacilityAddresses.City
            arrAddress(3) = oOwner.Facilities.FacilityAddresses.State
            arrAddress(4) = oOwner.Facilities.FacilityAddresses.Zip
            'strAddress = oOwner.Facilities.FacilityAddresses.AddressLine1 + "," + oOwner.Facilities.FacilityAddresses.AddressLine2 + "," + oOwner.Facilities.FacilityAddresses.City + "," + oOwner.Facilities.FacilityAddresses.State + "," + oOwner.Facilities.FacilityAddresses.Zip
            If oOwner.Facilities.FacilityAddresses.AddressId > 0 And oOwner.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtFacilityName.Text, arrAddress, "TEC", oOwner.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnTecnOwnerLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTecnOwnerLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = oOwner.Addresses.AddressLine1
            arrAddress(1) = oOwner.Addresses.AddressLine2
            arrAddress(2) = oOwner.Addresses.City
            arrAddress(3) = oOwner.Addresses.State
            arrAddress(4) = oOwner.Addresses.Zip
            'strAddress = oOwner.Addresses.AddressLine1 + "," + oOwner.Addresses.AddressLine2 + "," + oOwner.Addresses.City + "," + oOwner.Addresses.State + "," + oOwner.Addresses.Zip
            If oOwner.Addresses.AddressId > 0 And oOwner.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtOwnerName.Text, arrAddress, "TEC", oOwner.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnTecOwnerEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTecOwnerEnvelopes.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = oOwner.Addresses.AddressLine1
            arrAddress(1) = oOwner.Addresses.AddressLine2
            arrAddress(2) = oOwner.Addresses.City
            arrAddress(3) = oOwner.Addresses.State
            arrAddress(4) = oOwner.Addresses.Zip
            'strAddress = oOwner.Addresses.AddressLine1 + "," + oOwner.Addresses.AddressLine2 + "," + oOwner.Addresses.City + "," + oOwner.Addresses.State + "," + oOwner.Addresses.Zip
            If oOwner.Addresses.AddressId > 0 And oOwner.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(Me.txtOwnerName.Text, arrAddress, "TEC", oOwner.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

    Private Sub txtERAC_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtERAC.DoubleClick
        Try
            If MsgBox("Do you want to reset the ERAC to blank?", MsgBoxStyle.YesNo, "Reset ERAC") = MsgBoxResult.Yes Then
                oLustEvent.ERAC = 0
                txtERAC.Text = ""
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub txtIRAC_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIRAC.DoubleClick
        Try
            If MsgBox("Do you want to reset the IRAC to blank?", MsgBoxStyle.YesNo, "Reset IRAC") = MsgBoxResult.Yes Then
                oLustEvent.IRAC = 0
                txtIRAC.Text = ""
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
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

    Private Sub lblProjectManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblProjectManager.Click

    End Sub

    Private Sub btnActPlanning_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnActPlanning.Click
        Dim evtActPlanning As New ActivityPlanning(614, oLustEvent.FacilityID, oOwner.Facility.Name, oLustEvent.ID, oLustEvent.EVENTSEQUENCE)

        evtActPlanning.ShowDialog()

        evtActPlanning.Dispose()

    End Sub

    Private Sub txtERAC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtERAC.TextChanged
        oLustEvent.IsDirty = True

        btnSaveLUSTEvent.Enabled = True

    End Sub

    Private Sub txtIRAC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIRAC.TextChanged

        oLustEvent.IsDirty = True

        btnSaveLUSTEvent.Enabled = True

    End Sub

    Private Sub btnAddRemediationSystem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddRemediationSystem.Click
        Dim frmRemSysList As New RemediationSystemList
        Dim MyFrm As MusterContainer
        Try

            frmRemSysList.CallingForm = Me
            frmRemSysList.Mode = 0 ' Add
            frmRemSysList.EventActivityID = ugActivitiesandDocuments.ActiveRow.Cells("Event_Activity_ID").Value
            'MsgBox(frmRemSysList.EventActivityID)
            frmRemSysList.ShowDialog()
            GetLustRemediationForEvent(oLustEvent.ID)
            ' MyFrm.RefreshCalendarInfo()
            'MyFrm.LoadDueToMeCalendar()
            'MyFrm.LoadToDoCalendar()
            Me.SetupModifyLustEventForm(True, False)
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Add Remediation System" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmRemSysList = Nothing
        End Try
    End Sub

    Private Sub BtnSaveEngineers_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnSaveEngineers.Click
        'Update Record
        Try
            oLustEvent.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            MsgBox("Engineer/IRAC contact data has been saved successfully", MsgBoxStyle.OKOnly, "Financial View of Technical Data - Save ERAC/IRAC")
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Processing Save Engineers: " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

   
    
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCompAssDate.Click

    End Sub
End Class