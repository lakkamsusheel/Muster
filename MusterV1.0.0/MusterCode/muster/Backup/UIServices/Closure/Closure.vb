'-------------------------------------------------------------------------------
' MUSTER.MUSTER.MusterContainer.vb
'   Provides the parent MDI window for the application.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'1.2.0.4    KA          6/11/2007  Added logic to bring scroll bar on 
'                                  UGTanksAndPipes to last record modified.
''-------------------------------------------------------------------------------
Imports System.Data.SqlClient
Imports System.Text
Public Class Closure
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private WithEvents pOwn As MUSTER.BusinessLogic.pOwner
    Private WithEvents pClosure As New MUSTER.BusinessLogic.pClosureEvent
    'Private oEntity As New MUSTER.BusinessLogic.pEntity
    Public strFacilityIdTags As String
    Public nFacilityID As Integer
    Public MyGuid As New System.Guid
    Private bolLoading As Boolean = False
    Private bolCloEventTankPipeGridUpdateInProcess As Boolean = False
    Public bolNewPersona As Boolean = False
    Private BolAll_AAL_HasBackFillOnly As Boolean = False
    Private bolDisplayErrmessage As Boolean = True
    Private nErrMessage As Integer = 0
    Private dtMaxDate As DateTime = "1/1/1900"
    Private strMedia As String = " soil/ water"
    Private bolValidateSuccess As Boolean = True
    Private oAddressInfo As MUSTER.Info.AddressInfo
    'Private WithEvents AddrForm As GenAddressMSFT
    Private WithEvents AddressForm As Address
    Private dtNullDate As Object = Nothing
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    'Private chkCheckList As CheckBox
    Private dtAnalysis As New DataTable
    Private strActivetbPage As String
    'Private nSampleNumber As Integer = -1
    Private WithEvents objCntSearch As ContactSearch
    Private dsContacts As DataSet
    Private result As DialogResult
    Private alAnalysisType As New ArrayList
    Private strAnalysisType As String
    Private WithEvents SF As ShowFlags
    Private bolFrmActivated As Boolean = False
    Private bolFromNOI As Boolean = False
    Private strClosureEventIdTags As String
    Private slLetterCount As SortedList
    Dim returnVal As String = String.Empty
    Friend nCurrentEventID As Integer = -1
    Private pConStruct As New MUSTER.BusinessLogic.pContactStruct
    Private nLastTankPipeId As Integer = 0
    Private sLastTankPipeType As String = String.Empty
#Region "Company events and variables"
    Private WithEvents oCompanySearch As CompanySearch
    Dim pCompany As New MUSTER.BusinessLogic.pCompany
    Dim pLicensee As New MUSTER.BusinessLogic.pLicensee
    Dim strFromCompanySearch As String
#End Region
#End Region
#Region " Windows Form Designer generated code "
    Public Sub New(ByRef oOwner As MUSTER.BusinessLogic.pOwner, ByVal OwnerID As Int64, ByVal FacilityID As Int64, Optional ByVal nClosureID As Integer = 0)
        MyBase.New()
        MyGuid = System.Guid.NewGuid
        bolLoading = True
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        If oOwner.ID = 0 AndAlso OwnerID > 0 Then
            oOwner.Retrieve(OwnerID)
        End If
        pOwn = oOwner

        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppUser.LogEntry(Me.Text, MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Closure")
        Try



            InitControls()
            If OwnerID > 0 Then
                PopulateOwnerInfo(Integer.Parse(OwnerID))
                If FacilityID <= 0 Then
                    tbCntrlClosure.SelectedTab = tbPageOwnerDetail
                    Me.Text = "Closure - Owner Detail (" & txtOwnerName.Text & ")"
                End If
            End If

            If FacilityID > 0 Then
                tbCntrlClosure.SelectedTab = tbPageFacilityDetail
                PopulateFacilityInfo(Integer.Parse(FacilityID))
            End If

            If nClosureID > 0 Then
                Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
                Dim sender As Object
                Dim e As System.EventArgs
                For Each ugRow In dgClosureFacilityDetails.Rows
                    If ugRow.Cells("NOI ID").Value = nClosureID Then
                        dgClosureFacilityDetails.ActiveRow = ugRow
                        Me.SetupModifyClosureEventForm(nClosureID)
                    End If
                    Exit Sub
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub New(Optional ByVal OwnerID As Int64 = 0, Optional ByVal FacilityID As Int64 = 0)
        MyBase.New()
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppUser.LogEntry(Me.Text, MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Closure")
        Try
            InitControls()
            If OwnerID > 0 Then
                PopulateOwnerInfo(Integer.Parse(OwnerID))
                If FacilityID <= 0 Then
                    tbCntrlClosure.SelectedTab = tbPageOwnerDetail
                    Me.Text = "Closure - Owner Detail (" & txtOwnerName.Text & ")"
                End If
            End If

            If FacilityID > 0 Then
                tbCntrlClosure.SelectedTab = tbPageFacilityDetail
                PopulateFacilityInfo(Integer.Parse(FacilityID))
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub New(ByRef oOwner As MUSTER.BusinessLogic.pOwner)
        MyBase.New()
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        pOwn = oOwner

        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppUser.LogEntry(Me.Text, MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Closure")
        'Ends here 
        Try

            InitControls()
            PopulateOwnerInfo(pOwn.ID)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
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
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents tbPageOwnerDetail As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityDetail As System.Windows.Forms.TabPage
    Friend WithEvents lblProjectManager As System.Windows.Forms.Label
    Friend WithEvents tbPageSummary As System.Windows.Forms.TabPage
    Friend WithEvents Panel12 As System.Windows.Forms.Panel
    Friend WithEvents pnlLUSTEventBottom As System.Windows.Forms.Panel
    Friend WithEvents tbPageClosures As System.Windows.Forms.TabPage
    Friend WithEvents pnlNoticeOfInterestHeader As System.Windows.Forms.Panel
    Friend WithEvents lblClosureCountValue As System.Windows.Forms.Label
    Friend WithEvents lblClosureIDValue As System.Windows.Forms.Label
    Friend WithEvents lblClosureID As System.Windows.Forms.Label
    Friend WithEvents lblNOIReceived As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblStatusValue As System.Windows.Forms.Label
    Friend WithEvents btnReopen As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnNOIReceivedNo As System.Windows.Forms.RadioButton
    Friend WithEvents rbtnNOIReceivedYes As System.Windows.Forms.RadioButton
    Friend WithEvents pnlOwnerDetail As System.Windows.Forms.Panel
    Friend WithEvents lblNewOwnerSnippetValue As System.Windows.Forms.Label
    Friend WithEvents LinkLblCAPSignup As System.Windows.Forms.LinkLabel
    Friend WithEvents btnOwnerNameClose As System.Windows.Forms.Button
    Friend WithEvents btnOwnerNameCancel As System.Windows.Forms.Button
    Friend WithEvents btnOwnerNameOK As System.Windows.Forms.Button
    Friend WithEvents lblOwnerOrgName As System.Windows.Forms.Label
    Friend WithEvents btnOwnerNameSearch As System.Windows.Forms.Button
    Friend WithEvents lblOwnerNameSuffix As System.Windows.Forms.Label
    Friend WithEvents lblOwnerNameTitle As System.Windows.Forms.Label
    Friend WithEvents lblOwnerMiddleName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerLastName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerFirstName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerEmail As System.Windows.Forms.Label
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerFlag As System.Windows.Forms.Button
    Friend WithEvents btnOwnerComment As System.Windows.Forms.Button
    Friend WithEvents btnSaveOwner As System.Windows.Forms.Button
    Friend WithEvents btnOwnerCancel As System.Windows.Forms.Button
    Friend WithEvents lblOwnerAddress As System.Windows.Forms.Label
    Friend WithEvents lblOwnerName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerStatus As System.Windows.Forms.Label
    Friend WithEvents lblOwnerCapParticipant As System.Windows.Forms.Label
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Friend WithEvents lblOwnerType As System.Windows.Forms.Label
    Friend WithEvents lblOwnerID As System.Windows.Forms.Label
    Friend WithEvents lblOwnerPhone As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerBottom As System.Windows.Forms.Panel
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tbPageOwnerFacilities As System.Windows.Forms.TabPage
    Friend WithEvents lblUpcomingInstallDate As System.Windows.Forms.Label
    Friend WithEvents lnkLblNextFac As System.Windows.Forms.LinkLabel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnFacilityCancel As System.Windows.Forms.Button
    Friend WithEvents btnFacComments As System.Windows.Forms.Button
    Friend WithEvents btnFacilitySave As System.Windows.Forms.Button
    Friend WithEvents btnFacFlags As System.Windows.Forms.Button
    Friend WithEvents lblCAPStatus As System.Windows.Forms.Label
    Friend WithEvents ll As System.Windows.Forms.Label
    Friend WithEvents lnkLblPrevFacility As System.Windows.Forms.LinkLabel
    Friend WithEvents lblLUSTSite As System.Windows.Forms.Label
    Friend WithEvents lblPowerOff As System.Windows.Forms.Label
    Friend WithEvents lblCAPCandidate As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLocationType As System.Windows.Forms.Label
    Friend WithEvents lblFacilityMethod As System.Windows.Forms.Label
    Friend WithEvents lblFacilityDatum As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongMin As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongSec As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatMin As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatSec As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongDegree As System.Windows.Forms.Label
    Friend WithEvents lblFacilitySIC As System.Windows.Forms.Label
    Friend WithEvents txtFacilityFax As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityFax As System.Windows.Forms.Label
    Friend WithEvents lblDateReceived As System.Windows.Forms.Label
    Friend WithEvents txtFuelBrandcmb As System.Windows.Forms.ComboBox
    Friend WithEvents btnFacilityChangeCancel As System.Windows.Forms.Button
    Friend WithEvents lblPotentialOwner As System.Windows.Forms.Label
    Friend WithEvents lblFacilitySigOnNF As System.Windows.Forms.Label
    Friend WithEvents lblFacilityFuelBrand As System.Windows.Forms.Label
    Friend WithEvents lblFacilityStatus As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongitude As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatitude As System.Windows.Forms.Label
    Friend WithEvents lblFacilityType As System.Windows.Forms.Label
    Friend WithEvents lblfacilityAIID As System.Windows.Forms.Label
    Friend WithEvents lblFacilityID As System.Windows.Forms.Label
    Friend WithEvents txtfacilityPhone As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityPhone As System.Windows.Forms.Label
    Friend WithEvents lblFacilityAddress As System.Windows.Forms.Label
    Friend WithEvents lblFacilityName As System.Windows.Forms.Label
    Friend WithEvents txtFacilityZip As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityLatDegree As System.Windows.Forms.Label
    Friend WithEvents pnlFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents dgClosureFacilityDetails As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlClosureMain As System.Windows.Forms.Panel
    Friend WithEvents pnlClosuresDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlClosureReportDetails As System.Windows.Forms.Panel
    Friend WithEvents dtPickClosureReceived As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnProcessClosure As System.Windows.Forms.Button
    Friend WithEvents lblClosureReceived As System.Windows.Forms.Label
    Friend WithEvents lblClosureReportCertifiedContractor As System.Windows.Forms.Label
    Friend WithEvents lblClosureReportDateClosed As System.Windows.Forms.Label
    Friend WithEvents lblDateLastUsed As System.Windows.Forms.Label
    Friend WithEvents lblClosureReportCompany As System.Windows.Forms.Label
    Friend WithEvents lblClosureReportFillMaterial As System.Windows.Forms.Label
    Friend WithEvents dtPickDateClosed As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickDateLastUsed As System.Windows.Forms.DateTimePicker
    Friend WithEvents pnlClosureReport As System.Windows.Forms.Panel
    Friend WithEvents lblClosureReportHead As System.Windows.Forms.Label
    Friend WithEvents lblClosureReportDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlChecklistDetails As System.Windows.Forms.Panel
    Friend WithEvents lblDueBy As System.Windows.Forms.Label
    Friend WithEvents dtPickDueBy As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkClosures1 As System.Windows.Forms.CheckBox
    Friend WithEvents dtPickClosures1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateClosed As System.Windows.Forms.Label
    Friend WithEvents lblOpen As System.Windows.Forms.Label
    Friend WithEvents lblDateClosed1 As System.Windows.Forms.Label
    Friend WithEvents dtPickClosures2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures12 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures3 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures5 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures4 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures6 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures9 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures8 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures7 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures11 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures10 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures13 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures14 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures15 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures16 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures17 As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickClosures18 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlChecklist As System.Windows.Forms.Panel
    Friend WithEvents lblChecklistHead As System.Windows.Forms.Label
    Friend WithEvents lblChecklistDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlAnalysisDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlAnalysis As System.Windows.Forms.Panel
    Friend WithEvents lblAnalysisHead As System.Windows.Forms.Label
    Friend WithEvents lblAnalysisDisplay As System.Windows.Forms.Label
    Friend WithEvents PnlTanksPipesDetails As System.Windows.Forms.Panel
    Friend WithEvents ugTankandPipes As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlTanksPipes As System.Windows.Forms.Panel
    Friend WithEvents lblTanksPipesHead As System.Windows.Forms.Label
    Friend WithEvents lblTanksPipesDisplay As System.Windows.Forms.Label
    Friend WithEvents PnlNoticeOfInterestDetails As System.Windows.Forms.Panel
    Friend WithEvents cmbVerbalWaiver As System.Windows.Forms.CheckBox
    Friend WithEvents lblCertifiedContractor As System.Windows.Forms.Label
    Friend WithEvents cmbFillMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents lblFillMaterial As System.Windows.Forms.Label
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents lblScheduledDate As System.Windows.Forms.Label
    Friend WithEvents dtPickScheduledDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickReceived As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblReceived As System.Windows.Forms.Label
    Friend WithEvents btnProcessNOI As System.Windows.Forms.Button
    Friend WithEvents pnlNoticeOfInterest As System.Windows.Forms.Panel
    Friend WithEvents lblNoticeOfInterestHead As System.Windows.Forms.Label
    Friend WithEvents lblNoticeOfInterestDisplay As System.Windows.Forms.Label
    Public WithEvents tbCntrlClosure As System.Windows.Forms.TabControl
    Public WithEvents chkOwnerAgencyInterest As System.Windows.Forms.CheckBox
    Public WithEvents lblOwnerActiveOrNot As System.Windows.Forms.Label
    Public WithEvents lblCAPParticipationLevel As System.Windows.Forms.Label
    Public WithEvents txtOwnerOrgName As System.Windows.Forms.TextBox
    Public WithEvents rdOwnerOrg As System.Windows.Forms.RadioButton
    Public WithEvents rdOwnerPerson As System.Windows.Forms.RadioButton
    Public WithEvents cmbOwnerNameSuffix As System.Windows.Forms.ComboBox
    Public WithEvents cmbOwnerNameTitle As System.Windows.Forms.ComboBox
    Public WithEvents txtOwnerMiddleName As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerLastName As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerFirstName As System.Windows.Forms.TextBox
    Public WithEvents mskTxtOwnerFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone2 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone As AxMSMask.AxMaskEdBox
    Public WithEvents txtOwnerEmail As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerAddress As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerName As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerAIID As System.Windows.Forms.TextBox
    Public WithEvents lblOwnerIDValue As System.Windows.Forms.Label
    Public WithEvents cmbOwnerType As System.Windows.Forms.ComboBox
    Public WithEvents chkCAPParticipant As System.Windows.Forms.CheckBox
    Public WithEvents dtPickUpcomingInstallDateValue As System.Windows.Forms.DateTimePicker
    Public WithEvents chkUpcomingInstall As System.Windows.Forms.CheckBox
    Public WithEvents lblCAPStatusValue As System.Windows.Forms.Label
    Public WithEvents txtFuelBrand As System.Windows.Forms.TextBox
    Public WithEvents chkLUSTSite As System.Windows.Forms.CheckBox
    Public WithEvents chkCAPCandidate As System.Windows.Forms.CheckBox
    Public WithEvents cmbFacilityLocationType As System.Windows.Forms.ComboBox
    Public WithEvents cmbFacilityMethod As System.Windows.Forms.ComboBox
    Public WithEvents cmbFacilityDatum As System.Windows.Forms.ComboBox
    Public WithEvents cmbFacilityType As System.Windows.Forms.ComboBox
    Public WithEvents txtFacilityLatSec As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLongSec As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatMin As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLongMin As System.Windows.Forms.TextBox
    Public WithEvents mskTxtFacilityFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtFacilityPhone As AxMSMask.AxMaskEdBox
    Public WithEvents txtFacilityAddress As System.Windows.Forms.TextBox
    Public WithEvents dtPickFacilityRecvd As System.Windows.Forms.DateTimePicker
    Public WithEvents chkSignatureofNF As System.Windows.Forms.CheckBox
    Public WithEvents lblFacilityStatusValue As System.Windows.Forms.Label
    Public WithEvents txtFacilityLongDegree As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatDegree As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityAIID As System.Windows.Forms.TextBox
    Public WithEvents lblFacilityIDValue As System.Windows.Forms.Label
    Public WithEvents txtFacilityName As System.Windows.Forms.TextBox
    Public WithEvents ugFacilityList As Infragistics.Win.UltraWinGrid.UltraGrid
    Public WithEvents lblDateTransfered As System.Windows.Forms.Label
    Public WithEvents txtDueByNF As System.Windows.Forms.TextBox
    Public WithEvents pnlOwnerName As System.Windows.Forms.Panel
    Public WithEvents pnlOwnerOrg As System.Windows.Forms.Panel
    Public WithEvents pnlPersonOrganization As System.Windows.Forms.Panel
    Public WithEvents pnlOwnerPerson As System.Windows.Forms.Panel
    Public WithEvents pnlOwnerNameButton As System.Windows.Forms.Panel
    Friend WithEvents cmbClosureType As System.Windows.Forms.ComboBox
    Friend WithEvents btnAddClosure As System.Windows.Forms.Button
    Friend WithEvents pnlFacilityClosureButton As System.Windows.Forms.Panel
    Friend WithEvents lblNoOfClosuresValue As System.Windows.Forms.Label
    Friend WithEvents lblNoOfClosure As System.Windows.Forms.Label
    Friend WithEvents dGridAnalysis As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents chkShowAllTanksPipes As System.Windows.Forms.CheckBox
    Friend WithEvents btnEvtTankCollapse As System.Windows.Forms.Button
    Friend WithEvents udPreviousSubstance As Infragistics.Win.UltraWinGrid.UltraDropDown
    Friend WithEvents tbPageOwnerContactList As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerContactHeader As System.Windows.Forms.Panel
    Friend WithEvents chkOwnerShowActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents lblOwnerContacts As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerContactContainer As System.Windows.Forms.Panel
    Friend WithEvents ugOwnerContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerContactButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAddSearchContact As System.Windows.Forms.Button
    Public WithEvents pnl_FacilityDetail As System.Windows.Forms.Panel
    Friend WithEvents btnDeleteClosure As System.Windows.Forms.Button
    Friend WithEvents btnClosureComments As System.Windows.Forms.Button
    Friend WithEvents btnClosureCancel As System.Windows.Forms.Button
    Friend WithEvents btnSaveClosure As System.Windows.Forms.Button
    Friend WithEvents pnlOwnerFacilityBottom As System.Windows.Forms.Panel
    Public WithEvents lblNoOfFacilitiesValue As System.Windows.Forms.Label
    Friend WithEvents lblNoOfFacilities As System.Windows.Forms.Label
    Friend WithEvents tbCtrlFacClosureEvts As System.Windows.Forms.TabControl
    Friend WithEvents tbPageFacClosure As System.Windows.Forms.TabPage
    Friend WithEvents pnlContacts As System.Windows.Forms.Panel
    Friend WithEvents lblContactsHead As System.Windows.Forms.Label
    Friend WithEvents lblContactsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlContactDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlClosureContactHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlClosureContactsContainer As System.Windows.Forms.Panel
    Friend WithEvents ugClosureContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pnlClosureContactButtons As System.Windows.Forms.Panel
    Friend WithEvents chkClosureShowActive As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosureShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosureShowContactsForAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents lblClosureContacts As System.Windows.Forms.Label
    Friend WithEvents btnClosureContactModify As System.Windows.Forms.Button
    Friend WithEvents btnClosureContactDelete As System.Windows.Forms.Button
    Friend WithEvents btnClosureContactAssociate As System.Windows.Forms.Button
    Friend WithEvents btnClosureContactAddorSearch As System.Windows.Forms.Button
    Friend WithEvents cmbClosureReportFillMaterial As System.Windows.Forms.ComboBox
    Friend WithEvents txtClosureReportCompany As System.Windows.Forms.TextBox
    Friend WithEvents lblEnsite As System.Windows.Forms.Label
    Friend WithEvents txtLicensee As System.Windows.Forms.TextBox
    Friend WithEvents lblLicenseeSearch As System.Windows.Forms.Label
    Friend WithEvents lblNOISearchCompany As System.Windows.Forms.Label
    Friend WithEvents txtNOILicensee As System.Windows.Forms.TextBox
    Friend WithEvents txtNOICompany As System.Windows.Forms.TextBox
    Friend WithEvents dtPickClosures19 As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkClosures2 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures3 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures4 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures5 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures6 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures7 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures8 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures9 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures10 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures11 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures12 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures13 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures14 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures15 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures16 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures17 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures18 As System.Windows.Forms.CheckBox
    Friend WithEvents chkClosures19 As System.Windows.Forms.CheckBox
    Public WithEvents lblOwnerLastEditedOn As System.Windows.Forms.Label
    Public WithEvents lblOwnerLastEditedBy As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerSummaryHeader As System.Windows.Forms.Panel
    Public WithEvents UCOwnerSummary As MUSTER.OwnerSummary
    Friend WithEvents tbPageOwnerDocuments As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityDocuments As System.Windows.Forms.TabPage
    Friend WithEvents UCFacilityDocuments As MUSTER.DocumentViewControl
    Friend WithEvents UCOwnerDocuments As MUSTER.DocumentViewControl
    Friend WithEvents btnClosureOwnerLabels As System.Windows.Forms.Button
    Friend WithEvents btnClosureOwnerEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnClosureFacLabels As System.Windows.Forms.Button
    Friend WithEvents btnClosureFacEnvelopes As System.Windows.Forms.Button
    Public WithEvents txtFacilitySIC As System.Windows.Forms.Label
    Friend WithEvents lnkLblPrevClosure As System.Windows.Forms.LinkLabel
    Friend WithEvents lnkLblNextClosure As System.Windows.Forms.LinkLabel
    Friend WithEvents btnProcessNOIEnvelopes As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents dtPickAssess As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblMGPTF As System.Windows.Forms.Label
    Public WithEvents txtMGPTF As System.Windows.Forms.TextBox
    Friend WithEvents LblInspections As System.Windows.Forms.Label
    Friend WithEvents TxtInspections As System.Windows.Forms.TextBox
    Public WithEvents txtFeeBalance As System.Windows.Forms.TextBox
    Friend WithEvents LblFeeBalance As System.Windows.Forms.Label
    Public WithEvents dtFacilityPowerOff As System.Windows.Forms.DateTimePicker
    Public WithEvents lblInViolationValue As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Closure))
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.lblOwnerLastEditedOn = New System.Windows.Forms.Label
        Me.lblOwnerLastEditedBy = New System.Windows.Forms.Label
        Me.tbCntrlClosure = New System.Windows.Forms.TabControl
        Me.tbPageOwnerDetail = New System.Windows.Forms.TabPage
        Me.pnlOwnerBottom = New System.Windows.Forms.Panel
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.tbPageOwnerFacilities = New System.Windows.Forms.TabPage
        Me.ugFacilityList = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerFacilityBottom = New System.Windows.Forms.Panel
        Me.lblNoOfFacilitiesValue = New System.Windows.Forms.Label
        Me.lblNoOfFacilities = New System.Windows.Forms.Label
        Me.tbPageOwnerContactList = New System.Windows.Forms.TabPage
        Me.pnlOwnerContactContainer = New System.Windows.Forms.Panel
        Me.ugOwnerContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label2 = New System.Windows.Forms.Label
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
        Me.btnClosureOwnerLabels = New System.Windows.Forms.Button
        Me.btnClosureOwnerEnvelopes = New System.Windows.Forms.Button
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
        Me.btnSaveOwner = New System.Windows.Forms.Button
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
        Me.lblEnsite = New System.Windows.Forms.Label
        Me.lblOwnerIDValue = New System.Windows.Forms.Label
        Me.lblOwnerID = New System.Windows.Forms.Label
        Me.lblOwnerPhone = New System.Windows.Forms.Label
        Me.cmbOwnerType = New System.Windows.Forms.ComboBox
        Me.chkCAPParticipant = New System.Windows.Forms.CheckBox
        Me.tbPageFacilityDetail = New System.Windows.Forms.TabPage
        Me.pnlFacilityBottom = New System.Windows.Forms.Panel
        Me.tbCtrlFacClosureEvts = New System.Windows.Forms.TabControl
        Me.tbPageFacClosure = New System.Windows.Forms.TabPage
        Me.btnAddClosure = New System.Windows.Forms.Button
        Me.dgClosureFacilityDetails = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFacilityClosureButton = New System.Windows.Forms.Panel
        Me.lblNoOfClosuresValue = New System.Windows.Forms.Label
        Me.lblNoOfClosure = New System.Windows.Forms.Label
        Me.tbPageFacilityDocuments = New System.Windows.Forms.TabPage
        Me.UCFacilityDocuments = New MUSTER.DocumentViewControl
        Me.pnl_FacilityDetail = New System.Windows.Forms.Panel
        Me.txtFeeBalance = New System.Windows.Forms.TextBox
        Me.LblFeeBalance = New System.Windows.Forms.Label
        Me.txtMGPTF = New System.Windows.Forms.TextBox
        Me.lblMGPTF = New System.Windows.Forms.Label
        Me.dtPickAssess = New System.Windows.Forms.DateTimePicker
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnClosureFacLabels = New System.Windows.Forms.Button
        Me.btnClosureFacEnvelopes = New System.Windows.Forms.Button
        Me.dtPickUpcomingInstallDateValue = New System.Windows.Forms.DateTimePicker
        Me.lblUpcomingInstallDate = New System.Windows.Forms.Label
        Me.chkUpcomingInstall = New System.Windows.Forms.CheckBox
        Me.lnkLblNextFac = New System.Windows.Forms.LinkLabel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnFacilityCancel = New System.Windows.Forms.Button
        Me.btnFacComments = New System.Windows.Forms.Button
        Me.btnFacilitySave = New System.Windows.Forms.Button
        Me.btnFacFlags = New System.Windows.Forms.Button
        Me.lblCAPStatusValue = New System.Windows.Forms.Label
        Me.lblCAPStatus = New System.Windows.Forms.Label
        Me.txtFuelBrand = New System.Windows.Forms.TextBox
        Me.ll = New System.Windows.Forms.Label
        Me.dtFacilityPowerOff = New System.Windows.Forms.DateTimePicker
        Me.lnkLblPrevFacility = New System.Windows.Forms.LinkLabel
        Me.lblDateTransfered = New System.Windows.Forms.Label
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
        Me.txtFacilityFax = New System.Windows.Forms.TextBox
        Me.lblFacilityFax = New System.Windows.Forms.Label
        Me.dtPickFacilityRecvd = New System.Windows.Forms.DateTimePicker
        Me.lblDateReceived = New System.Windows.Forms.Label
        Me.txtFuelBrandcmb = New System.Windows.Forms.ComboBox
        Me.btnFacilityChangeCancel = New System.Windows.Forms.Button
        Me.txtDueByNF = New System.Windows.Forms.TextBox
        Me.chkSignatureofNF = New System.Windows.Forms.CheckBox
        Me.lblPotentialOwner = New System.Windows.Forms.Label
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
        Me.tbPageClosures = New System.Windows.Forms.TabPage
        Me.pnlClosureMain = New System.Windows.Forms.Panel
        Me.pnlClosuresDetails = New System.Windows.Forms.Panel
        Me.pnlContactDetails = New System.Windows.Forms.Panel
        Me.pnlClosureContactButtons = New System.Windows.Forms.Panel
        Me.btnClosureContactModify = New System.Windows.Forms.Button
        Me.btnClosureContactDelete = New System.Windows.Forms.Button
        Me.btnClosureContactAssociate = New System.Windows.Forms.Button
        Me.btnClosureContactAddorSearch = New System.Windows.Forms.Button
        Me.pnlClosureContactsContainer = New System.Windows.Forms.Panel
        Me.ugClosureContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label4 = New System.Windows.Forms.Label
        Me.pnlClosureContactHeader = New System.Windows.Forms.Panel
        Me.chkClosureShowActive = New System.Windows.Forms.CheckBox
        Me.chkClosureShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkClosureShowContactsForAllModules = New System.Windows.Forms.CheckBox
        Me.lblClosureContacts = New System.Windows.Forms.Label
        Me.pnlContacts = New System.Windows.Forms.Panel
        Me.lblContactsHead = New System.Windows.Forms.Label
        Me.lblContactsDisplay = New System.Windows.Forms.Label
        Me.pnlClosureReportDetails = New System.Windows.Forms.Panel
        Me.TxtInspections = New System.Windows.Forms.TextBox
        Me.LblInspections = New System.Windows.Forms.Label
        Me.lblLicenseeSearch = New System.Windows.Forms.Label
        Me.txtLicensee = New System.Windows.Forms.TextBox
        Me.txtClosureReportCompany = New System.Windows.Forms.TextBox
        Me.dtPickClosureReceived = New System.Windows.Forms.DateTimePicker
        Me.btnProcessClosure = New System.Windows.Forms.Button
        Me.lblClosureReceived = New System.Windows.Forms.Label
        Me.cmbClosureReportFillMaterial = New System.Windows.Forms.ComboBox
        Me.lblClosureReportCertifiedContractor = New System.Windows.Forms.Label
        Me.lblClosureReportDateClosed = New System.Windows.Forms.Label
        Me.lblDateLastUsed = New System.Windows.Forms.Label
        Me.lblClosureReportCompany = New System.Windows.Forms.Label
        Me.lblClosureReportFillMaterial = New System.Windows.Forms.Label
        Me.dtPickDateClosed = New System.Windows.Forms.DateTimePicker
        Me.dtPickDateLastUsed = New System.Windows.Forms.DateTimePicker
        Me.pnlClosureReport = New System.Windows.Forms.Panel
        Me.lblClosureReportHead = New System.Windows.Forms.Label
        Me.lblClosureReportDisplay = New System.Windows.Forms.Label
        Me.pnlChecklistDetails = New System.Windows.Forms.Panel
        Me.lblDueBy = New System.Windows.Forms.Label
        Me.dtPickDueBy = New System.Windows.Forms.DateTimePicker
        Me.chkClosures1 = New System.Windows.Forms.CheckBox
        Me.dtPickClosures1 = New System.Windows.Forms.DateTimePicker
        Me.lblDateClosed = New System.Windows.Forms.Label
        Me.lblOpen = New System.Windows.Forms.Label
        Me.lblDateClosed1 = New System.Windows.Forms.Label
        Me.dtPickClosures2 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures12 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures3 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures5 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures4 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures6 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures9 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures8 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures7 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures11 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures10 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures13 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures14 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures15 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures16 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures17 = New System.Windows.Forms.DateTimePicker
        Me.dtPickClosures18 = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.dtPickClosures19 = New System.Windows.Forms.DateTimePicker
        Me.chkClosures2 = New System.Windows.Forms.CheckBox
        Me.chkClosures3 = New System.Windows.Forms.CheckBox
        Me.chkClosures4 = New System.Windows.Forms.CheckBox
        Me.chkClosures5 = New System.Windows.Forms.CheckBox
        Me.chkClosures6 = New System.Windows.Forms.CheckBox
        Me.chkClosures7 = New System.Windows.Forms.CheckBox
        Me.chkClosures8 = New System.Windows.Forms.CheckBox
        Me.chkClosures9 = New System.Windows.Forms.CheckBox
        Me.chkClosures10 = New System.Windows.Forms.CheckBox
        Me.chkClosures11 = New System.Windows.Forms.CheckBox
        Me.chkClosures12 = New System.Windows.Forms.CheckBox
        Me.chkClosures13 = New System.Windows.Forms.CheckBox
        Me.chkClosures14 = New System.Windows.Forms.CheckBox
        Me.chkClosures15 = New System.Windows.Forms.CheckBox
        Me.chkClosures16 = New System.Windows.Forms.CheckBox
        Me.chkClosures17 = New System.Windows.Forms.CheckBox
        Me.chkClosures18 = New System.Windows.Forms.CheckBox
        Me.chkClosures19 = New System.Windows.Forms.CheckBox
        Me.pnlChecklist = New System.Windows.Forms.Panel
        Me.lblChecklistHead = New System.Windows.Forms.Label
        Me.lblChecklistDisplay = New System.Windows.Forms.Label
        Me.pnlAnalysisDetails = New System.Windows.Forms.Panel
        Me.dGridAnalysis = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlAnalysis = New System.Windows.Forms.Panel
        Me.lblAnalysisHead = New System.Windows.Forms.Label
        Me.lblAnalysisDisplay = New System.Windows.Forms.Label
        Me.PnlTanksPipesDetails = New System.Windows.Forms.Panel
        Me.udPreviousSubstance = New Infragistics.Win.UltraWinGrid.UltraDropDown
        Me.btnEvtTankCollapse = New System.Windows.Forms.Button
        Me.chkShowAllTanksPipes = New System.Windows.Forms.CheckBox
        Me.ugTankandPipes = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTanksPipes = New System.Windows.Forms.Panel
        Me.lblTanksPipesHead = New System.Windows.Forms.Label
        Me.lblTanksPipesDisplay = New System.Windows.Forms.Label
        Me.PnlNoticeOfInterestDetails = New System.Windows.Forms.Panel
        Me.lblNOISearchCompany = New System.Windows.Forms.Label
        Me.txtNOILicensee = New System.Windows.Forms.TextBox
        Me.txtNOICompany = New System.Windows.Forms.TextBox
        Me.cmbVerbalWaiver = New System.Windows.Forms.CheckBox
        Me.lblCertifiedContractor = New System.Windows.Forms.Label
        Me.cmbFillMaterial = New System.Windows.Forms.ComboBox
        Me.lblFillMaterial = New System.Windows.Forms.Label
        Me.lblCompany = New System.Windows.Forms.Label
        Me.lblScheduledDate = New System.Windows.Forms.Label
        Me.dtPickScheduledDate = New System.Windows.Forms.DateTimePicker
        Me.dtPickReceived = New System.Windows.Forms.DateTimePicker
        Me.lblReceived = New System.Windows.Forms.Label
        Me.btnProcessNOI = New System.Windows.Forms.Button
        Me.btnProcessNOIEnvelopes = New System.Windows.Forms.Button
        Me.pnlNoticeOfInterest = New System.Windows.Forms.Panel
        Me.lblNoticeOfInterestHead = New System.Windows.Forms.Label
        Me.lblNoticeOfInterestDisplay = New System.Windows.Forms.Label
        Me.pnlLUSTEventBottom = New System.Windows.Forms.Panel
        Me.btnDeleteClosure = New System.Windows.Forms.Button
        Me.btnClosureComments = New System.Windows.Forms.Button
        Me.btnClosureCancel = New System.Windows.Forms.Button
        Me.btnSaveClosure = New System.Windows.Forms.Button
        Me.pnlNoticeOfInterestHeader = New System.Windows.Forms.Panel
        Me.lnkLblNextClosure = New System.Windows.Forms.LinkLabel
        Me.lnkLblPrevClosure = New System.Windows.Forms.LinkLabel
        Me.btnReopen = New System.Windows.Forms.Button
        Me.lblNOIReceived = New System.Windows.Forms.Label
        Me.lblClosureCountValue = New System.Windows.Forms.Label
        Me.lblClosureIDValue = New System.Windows.Forms.Label
        Me.lblClosureID = New System.Windows.Forms.Label
        Me.cmbClosureType = New System.Windows.Forms.ComboBox
        Me.lblProjectManager = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.lblStatusValue = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rbtnNOIReceivedNo = New System.Windows.Forms.RadioButton
        Me.rbtnNOIReceivedYes = New System.Windows.Forms.RadioButton
        Me.tbPageSummary = New System.Windows.Forms.TabPage
        Me.UCOwnerSummary = New MUSTER.OwnerSummary
        Me.pnlOwnerSummaryHeader = New System.Windows.Forms.Panel
        Me.Panel12 = New System.Windows.Forms.Panel
        Me.lblInViolationValue = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbCntrlClosure.SuspendLayout()
        Me.tbPageOwnerDetail.SuspendLayout()
        Me.pnlOwnerBottom.SuspendLayout()
        Me.TabControl1.SuspendLayout()
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
        Me.tbCtrlFacClosureEvts.SuspendLayout()
        Me.tbPageFacClosure.SuspendLayout()
        CType(Me.dgClosureFacilityDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityClosureButton.SuspendLayout()
        Me.tbPageFacilityDocuments.SuspendLayout()
        Me.pnl_FacilityDetail.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageClosures.SuspendLayout()
        Me.pnlClosureMain.SuspendLayout()
        Me.pnlClosuresDetails.SuspendLayout()
        Me.pnlContactDetails.SuspendLayout()
        Me.pnlClosureContactButtons.SuspendLayout()
        Me.pnlClosureContactsContainer.SuspendLayout()
        CType(Me.ugClosureContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlClosureContactHeader.SuspendLayout()
        Me.pnlContacts.SuspendLayout()
        Me.pnlClosureReportDetails.SuspendLayout()
        Me.pnlClosureReport.SuspendLayout()
        Me.pnlChecklistDetails.SuspendLayout()
        Me.pnlChecklist.SuspendLayout()
        Me.pnlAnalysisDetails.SuspendLayout()
        CType(Me.dGridAnalysis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAnalysis.SuspendLayout()
        Me.PnlTanksPipesDetails.SuspendLayout()
        CType(Me.udPreviousSubstance, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugTankandPipes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTanksPipes.SuspendLayout()
        Me.PnlNoticeOfInterestDetails.SuspendLayout()
        Me.pnlNoticeOfInterest.SuspendLayout()
        Me.pnlLUSTEventBottom.SuspendLayout()
        Me.pnlNoticeOfInterestHeader.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.tbPageSummary.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.TabIndex = 3
        '
        'lblOwnerLastEditedOn
        '
        Me.lblOwnerLastEditedOn.Location = New System.Drawing.Point(664, 8)
        Me.lblOwnerLastEditedOn.Name = "lblOwnerLastEditedOn"
        Me.lblOwnerLastEditedOn.Size = New System.Drawing.Size(184, 16)
        Me.lblOwnerLastEditedOn.TabIndex = 0
        '
        'lblOwnerLastEditedBy
        '
        Me.lblOwnerLastEditedBy.Location = New System.Drawing.Point(345, 8)
        Me.lblOwnerLastEditedBy.Name = "lblOwnerLastEditedBy"
        Me.lblOwnerLastEditedBy.Size = New System.Drawing.Size(168, 16)
        Me.lblOwnerLastEditedBy.TabIndex = 1017
        Me.lblOwnerLastEditedBy.Text = "Last Edited By :"
        '
        'tbCntrlClosure
        '
        Me.tbCntrlClosure.Controls.Add(Me.tbPageOwnerDetail)
        Me.tbCntrlClosure.Controls.Add(Me.tbPageFacilityDetail)
        Me.tbCntrlClosure.Controls.Add(Me.tbPageClosures)
        Me.tbCntrlClosure.Controls.Add(Me.tbPageSummary)
        Me.tbCntrlClosure.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCntrlClosure.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCntrlClosure.ItemSize = New System.Drawing.Size(64, 18)
        Me.tbCntrlClosure.Location = New System.Drawing.Point(0, 0)
        Me.tbCntrlClosure.Multiline = True
        Me.tbCntrlClosure.Name = "tbCntrlClosure"
        Me.tbCntrlClosure.SelectedIndex = 0
        Me.tbCntrlClosure.ShowToolTips = True
        Me.tbCntrlClosure.Size = New System.Drawing.Size(972, 694)
        Me.tbCntrlClosure.TabIndex = 2
        '
        'tbPageOwnerDetail
        '
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerBottom)
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerDetail)
        Me.tbPageOwnerDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageOwnerDetail.Name = "tbPageOwnerDetail"
        Me.tbPageOwnerDetail.Size = New System.Drawing.Size(964, 668)
        Me.tbPageOwnerDetail.TabIndex = 7
        Me.tbPageOwnerDetail.Text = "Owner Details"
        '
        'pnlOwnerBottom
        '
        Me.pnlOwnerBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlOwnerBottom.Controls.Add(Me.TabControl1)
        Me.pnlOwnerBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerBottom.Location = New System.Drawing.Point(0, 264)
        Me.pnlOwnerBottom.Name = "pnlOwnerBottom"
        Me.pnlOwnerBottom.Size = New System.Drawing.Size(964, 404)
        Me.pnlOwnerBottom.TabIndex = 45
        '
        'TabControl1
        '
        Me.TabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.TabControl1.Controls.Add(Me.tbPageOwnerFacilities)
        Me.TabControl1.Controls.Add(Me.tbPageOwnerContactList)
        Me.TabControl1.Controls.Add(Me.tbPageOwnerDocuments)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(962, 402)
        Me.TabControl1.TabIndex = 0
        '
        'tbPageOwnerFacilities
        '
        Me.tbPageOwnerFacilities.Controls.Add(Me.ugFacilityList)
        Me.tbPageOwnerFacilities.Controls.Add(Me.pnlOwnerFacilityBottom)
        Me.tbPageOwnerFacilities.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerFacilities.Name = "tbPageOwnerFacilities"
        Me.tbPageOwnerFacilities.Size = New System.Drawing.Size(954, 371)
        Me.tbPageOwnerFacilities.TabIndex = 0
        Me.tbPageOwnerFacilities.Text = "Facilities"
        '
        'ugFacilityList
        '
        Me.ugFacilityList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilityList.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFacilityList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFacilityList.Location = New System.Drawing.Point(0, 0)
        Me.ugFacilityList.Name = "ugFacilityList"
        Me.ugFacilityList.Size = New System.Drawing.Size(954, 347)
        Me.ugFacilityList.TabIndex = 3
        '
        'pnlOwnerFacilityBottom
        '
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.lblNoOfFacilitiesValue)
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.lblNoOfFacilities)
        Me.pnlOwnerFacilityBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerFacilityBottom.Location = New System.Drawing.Point(0, 347)
        Me.pnlOwnerFacilityBottom.Name = "pnlOwnerFacilityBottom"
        Me.pnlOwnerFacilityBottom.Size = New System.Drawing.Size(954, 24)
        Me.pnlOwnerFacilityBottom.TabIndex = 22
        '
        'lblNoOfFacilitiesValue
        '
        Me.lblNoOfFacilitiesValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilitiesValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfFacilitiesValue.Location = New System.Drawing.Point(100, 0)
        Me.lblNoOfFacilitiesValue.Name = "lblNoOfFacilitiesValue"
        Me.lblNoOfFacilitiesValue.Size = New System.Drawing.Size(56, 24)
        Me.lblNoOfFacilitiesValue.TabIndex = 1
        Me.lblNoOfFacilitiesValue.Text = "0"
        '
        'lblNoOfFacilities
        '
        Me.lblNoOfFacilities.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilities.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfFacilities.Location = New System.Drawing.Point(0, 0)
        Me.lblNoOfFacilities.Name = "lblNoOfFacilities"
        Me.lblNoOfFacilities.Size = New System.Drawing.Size(100, 24)
        Me.lblNoOfFacilities.TabIndex = 0
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
        Me.tbPageOwnerContactList.Size = New System.Drawing.Size(954, 371)
        Me.tbPageOwnerContactList.TabIndex = 1
        Me.tbPageOwnerContactList.Text = "Contacts"
        '
        'pnlOwnerContactContainer
        '
        Me.pnlOwnerContactContainer.Controls.Add(Me.ugOwnerContacts)
        Me.pnlOwnerContactContainer.Controls.Add(Me.Label2)
        Me.pnlOwnerContactContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerContactContainer.Location = New System.Drawing.Point(0, 25)
        Me.pnlOwnerContactContainer.Name = "pnlOwnerContactContainer"
        Me.pnlOwnerContactContainer.Size = New System.Drawing.Size(954, 316)
        Me.pnlOwnerContactContainer.TabIndex = 3
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
        Me.ugOwnerContacts.Size = New System.Drawing.Size(954, 316)
        Me.ugOwnerContacts.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(792, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(7, 23)
        Me.Label2.TabIndex = 2
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
        Me.pnlOwnerContactHeader.Size = New System.Drawing.Size(954, 25)
        Me.pnlOwnerContactHeader.TabIndex = 1
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
        Me.lblOwnerContacts.Size = New System.Drawing.Size(128, 16)
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
        Me.pnlOwnerContactButtons.Location = New System.Drawing.Point(0, 341)
        Me.pnlOwnerContactButtons.Name = "pnlOwnerContactButtons"
        Me.pnlOwnerContactButtons.Size = New System.Drawing.Size(954, 30)
        Me.pnlOwnerContactButtons.TabIndex = 4
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
        Me.tbPageOwnerDocuments.Size = New System.Drawing.Size(954, 371)
        Me.tbPageOwnerDocuments.TabIndex = 2
        Me.tbPageOwnerDocuments.Text = "Documents"
        '
        'UCOwnerDocuments
        '
        Me.UCOwnerDocuments.AutoScroll = True
        Me.UCOwnerDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerDocuments.Name = "UCOwnerDocuments"
        Me.UCOwnerDocuments.Size = New System.Drawing.Size(954, 371)
        Me.UCOwnerDocuments.TabIndex = 2
        '
        'pnlOwnerDetail
        '
        Me.pnlOwnerDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnlOwnerDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlOwnerDetail.Controls.Add(Me.btnClosureOwnerLabels)
        Me.pnlOwnerDetail.Controls.Add(Me.btnClosureOwnerEnvelopes)
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
        Me.pnlOwnerDetail.Controls.Add(Me.lblEnsite)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerIDValue)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerID)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerPhone)
        Me.pnlOwnerDetail.Controls.Add(Me.cmbOwnerType)
        Me.pnlOwnerDetail.Controls.Add(Me.chkCAPParticipant)
        Me.pnlOwnerDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerDetail.Name = "pnlOwnerDetail"
        Me.pnlOwnerDetail.Size = New System.Drawing.Size(964, 264)
        Me.pnlOwnerDetail.TabIndex = 45
        '
        'btnClosureOwnerLabels
        '
        Me.btnClosureOwnerLabels.Location = New System.Drawing.Point(4, 116)
        Me.btnClosureOwnerLabels.Name = "btnClosureOwnerLabels"
        Me.btnClosureOwnerLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnClosureOwnerLabels.TabIndex = 1070
        Me.btnClosureOwnerLabels.Text = "Labels"
        '
        'btnClosureOwnerEnvelopes
        '
        Me.btnClosureOwnerEnvelopes.Location = New System.Drawing.Point(4, 88)
        Me.btnClosureOwnerEnvelopes.Name = "btnClosureOwnerEnvelopes"
        Me.btnClosureOwnerEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnClosureOwnerEnvelopes.TabIndex = 1069
        Me.btnClosureOwnerEnvelopes.Text = "Envelopes"
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
        Me.chkOwnerAgencyInterest.Location = New System.Drawing.Point(552, 16)
        Me.chkOwnerAgencyInterest.Name = "chkOwnerAgencyInterest"
        Me.chkOwnerAgencyInterest.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkOwnerAgencyInterest.Size = New System.Drawing.Size(112, 20)
        Me.chkOwnerAgencyInterest.TabIndex = 1008
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
        Me.LinkLblCAPSignup.Enabled = False
        Me.LinkLblCAPSignup.Location = New System.Drawing.Point(552, 74)
        Me.LinkLblCAPSignup.Name = "LinkLblCAPSignup"
        Me.LinkLblCAPSignup.Size = New System.Drawing.Size(152, 16)
        Me.LinkLblCAPSignup.TabIndex = 1005
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
        Me.pnlOwnerName.Location = New System.Drawing.Point(24, 224)
        Me.pnlOwnerName.Name = "pnlOwnerName"
        Me.pnlOwnerName.Size = New System.Drawing.Size(296, 256)
        Me.pnlOwnerName.TabIndex = 3
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
        Me.btnOwnerNameClose.Size = New System.Drawing.Size(56, 19)
        Me.btnOwnerNameClose.TabIndex = 2
        Me.btnOwnerNameClose.Text = "Close"
        '
        'btnOwnerNameCancel
        '
        Me.btnOwnerNameCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerNameCancel.Enabled = False
        Me.btnOwnerNameCancel.Location = New System.Drawing.Point(104, 0)
        Me.btnOwnerNameCancel.Name = "btnOwnerNameCancel"
        Me.btnOwnerNameCancel.Size = New System.Drawing.Size(56, 19)
        Me.btnOwnerNameCancel.TabIndex = 1
        Me.btnOwnerNameCancel.Text = "Cancel"
        '
        'btnOwnerNameOK
        '
        Me.btnOwnerNameOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerNameOK.Location = New System.Drawing.Point(56, 1)
        Me.btnOwnerNameOK.Name = "btnOwnerNameOK"
        Me.btnOwnerNameOK.Size = New System.Drawing.Size(41, 19)
        Me.btnOwnerNameOK.TabIndex = 0
        Me.btnOwnerNameOK.Text = "Save"
        '
        'pnlOwnerOrg
        '
        Me.pnlOwnerOrg.Controls.Add(Me.txtOwnerOrgName)
        Me.pnlOwnerOrg.Controls.Add(Me.lblOwnerOrgName)
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
        Me.mskTxtOwnerFax.TabIndex = 6
        '
        'mskTxtOwnerPhone2
        '
        Me.mskTxtOwnerPhone2.ContainingControl = Me
        Me.mskTxtOwnerPhone2.Location = New System.Drawing.Point(424, 88)
        Me.mskTxtOwnerPhone2.Name = "mskTxtOwnerPhone2"
        Me.mskTxtOwnerPhone2.OcxState = CType(resources.GetObject("mskTxtOwnerPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone2.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtOwnerPhone2.TabIndex = 5
        '
        'mskTxtOwnerPhone
        '
        Me.mskTxtOwnerPhone.ContainingControl = Me
        Me.mskTxtOwnerPhone.Location = New System.Drawing.Point(424, 64)
        Me.mskTxtOwnerPhone.Name = "mskTxtOwnerPhone"
        Me.mskTxtOwnerPhone.OcxState = CType(resources.GetObject("mskTxtOwnerPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtOwnerPhone.TabIndex = 4
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
        Me.txtOwnerEmail.TabIndex = 12
        Me.txtOwnerEmail.Text = ""
        Me.txtOwnerEmail.WordWrap = False
        '
        'lblFax
        '
        Me.lblFax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFax.Location = New System.Drawing.Point(344, 112)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(80, 20)
        Me.lblFax.TabIndex = 44
        Me.lblFax.Text = "Fax:"
        Me.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlOwnerButtons
        '
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerFlag)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerComment)
        Me.pnlOwnerButtons.Controls.Add(Me.btnSaveOwner)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerCancel)
        Me.pnlOwnerButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.pnlOwnerButtons.Location = New System.Drawing.Point(336, 176)
        Me.pnlOwnerButtons.Name = "pnlOwnerButtons"
        Me.pnlOwnerButtons.Size = New System.Drawing.Size(335, 40)
        Me.pnlOwnerButtons.TabIndex = 57
        '
        'btnOwnerFlag
        '
        Me.btnOwnerFlag.Location = New System.Drawing.Point(158, 8)
        Me.btnOwnerFlag.Name = "btnOwnerFlag"
        Me.btnOwnerFlag.Size = New System.Drawing.Size(75, 26)
        Me.btnOwnerFlag.TabIndex = 48
        Me.btnOwnerFlag.Text = "Flags"
        '
        'btnOwnerComment
        '
        Me.btnOwnerComment.Location = New System.Drawing.Point(238, 8)
        Me.btnOwnerComment.Name = "btnOwnerComment"
        Me.btnOwnerComment.Size = New System.Drawing.Size(80, 26)
        Me.btnOwnerComment.TabIndex = 47
        Me.btnOwnerComment.Text = "Comments"
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
        'btnOwnerCancel
        '
        Me.btnOwnerCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerCancel.Enabled = False
        Me.btnOwnerCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerCancel.Location = New System.Drawing.Point(100, 8)
        Me.btnOwnerCancel.Name = "btnOwnerCancel"
        Me.btnOwnerCancel.Size = New System.Drawing.Size(54, 26)
        Me.btnOwnerCancel.TabIndex = 42
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
        Me.txtOwnerAddress.TabIndex = 3
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
        Me.txtOwnerName.TabIndex = 2
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
        Me.lblOwnerStatus.Location = New System.Drawing.Point(344, 16)
        Me.lblOwnerStatus.Name = "lblOwnerStatus"
        Me.lblOwnerStatus.Size = New System.Drawing.Size(80, 20)
        Me.lblOwnerStatus.TabIndex = 84
        Me.lblOwnerStatus.Text = "Owner Status:"
        Me.lblOwnerStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOwnerCapParticipant
        '
        Me.lblOwnerCapParticipant.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerCapParticipant.Location = New System.Drawing.Point(552, 48)
        Me.lblOwnerCapParticipant.Name = "lblOwnerCapParticipant"
        Me.lblOwnerCapParticipant.Size = New System.Drawing.Size(128, 20)
        Me.lblOwnerCapParticipant.TabIndex = 52
        Me.lblOwnerCapParticipant.Text = "CAP Participation Level:"
        Me.lblOwnerCapParticipant.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPhone2
        '
        Me.lblPhone2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhone2.Location = New System.Drawing.Point(344, 88)
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
        Me.lblOwnerType.Size = New System.Drawing.Size(70, 20)
        Me.lblOwnerType.TabIndex = 40
        Me.lblOwnerType.Text = "Owner Type:"
        Me.lblOwnerType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtOwnerAIID
        '
        Me.txtOwnerAIID.AcceptsTab = True
        Me.txtOwnerAIID.AutoSize = False
        Me.txtOwnerAIID.Enabled = False
        Me.txtOwnerAIID.Location = New System.Drawing.Point(424, 40)
        Me.txtOwnerAIID.Name = "txtOwnerAIID"
        Me.txtOwnerAIID.Size = New System.Drawing.Size(96, 20)
        Me.txtOwnerAIID.TabIndex = 7
        Me.txtOwnerAIID.Text = ""
        Me.txtOwnerAIID.WordWrap = False
        '
        'lblEnsite
        '
        Me.lblEnsite.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEnsite.Location = New System.Drawing.Point(344, 40)
        Me.lblEnsite.Name = "lblEnsite"
        Me.lblEnsite.Size = New System.Drawing.Size(80, 20)
        Me.lblEnsite.TabIndex = 38
        Me.lblEnsite.Text = "Ensite ID:"
        Me.lblEnsite.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOwnerIDValue
        '
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
        Me.lblOwnerPhone.Location = New System.Drawing.Point(344, 64)
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
        Me.cmbOwnerType.ItemHeight = 15
        Me.cmbOwnerType.Location = New System.Drawing.Point(80, 160)
        Me.cmbOwnerType.Name = "cmbOwnerType"
        Me.cmbOwnerType.Size = New System.Drawing.Size(248, 23)
        Me.cmbOwnerType.TabIndex = 1
        Me.cmbOwnerType.ValueMember = "1"
        '
        'chkCAPParticipant
        '
        Me.chkCAPParticipant.Checked = True
        Me.chkCAPParticipant.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCAPParticipant.Location = New System.Drawing.Point(824, 16)
        Me.chkCAPParticipant.Name = "chkCAPParticipant"
        Me.chkCAPParticipant.Size = New System.Drawing.Size(16, 20)
        Me.chkCAPParticipant.TabIndex = 28
        Me.chkCAPParticipant.Visible = False
        '
        'tbPageFacilityDetail
        '
        Me.tbPageFacilityDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageFacilityDetail.Controls.Add(Me.pnlFacilityBottom)
        Me.tbPageFacilityDetail.Controls.Add(Me.pnl_FacilityDetail)
        Me.tbPageFacilityDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageFacilityDetail.Name = "tbPageFacilityDetail"
        Me.tbPageFacilityDetail.Size = New System.Drawing.Size(964, 668)
        Me.tbPageFacilityDetail.TabIndex = 8
        Me.tbPageFacilityDetail.Text = "Facility Details"
        Me.tbPageFacilityDetail.Visible = False
        '
        'pnlFacilityBottom
        '
        Me.pnlFacilityBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlFacilityBottom.Controls.Add(Me.tbCtrlFacClosureEvts)
        Me.pnlFacilityBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityBottom.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlFacilityBottom.Location = New System.Drawing.Point(0, 288)
        Me.pnlFacilityBottom.Name = "pnlFacilityBottom"
        Me.pnlFacilityBottom.Size = New System.Drawing.Size(960, 376)
        Me.pnlFacilityBottom.TabIndex = 99
        '
        'tbCtrlFacClosureEvts
        '
        Me.tbCtrlFacClosureEvts.Controls.Add(Me.tbPageFacClosure)
        Me.tbCtrlFacClosureEvts.Controls.Add(Me.tbPageFacilityDocuments)
        Me.tbCtrlFacClosureEvts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlFacClosureEvts.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlFacClosureEvts.Name = "tbCtrlFacClosureEvts"
        Me.tbCtrlFacClosureEvts.SelectedIndex = 0
        Me.tbCtrlFacClosureEvts.Size = New System.Drawing.Size(958, 374)
        Me.tbCtrlFacClosureEvts.TabIndex = 100
        '
        'tbPageFacClosure
        '
        Me.tbPageFacClosure.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageFacClosure.Controls.Add(Me.btnAddClosure)
        Me.tbPageFacClosure.Controls.Add(Me.dgClosureFacilityDetails)
        Me.tbPageFacClosure.Controls.Add(Me.pnlFacilityClosureButton)
        Me.tbPageFacClosure.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacClosure.Name = "tbPageFacClosure"
        Me.tbPageFacClosure.Size = New System.Drawing.Size(950, 346)
        Me.tbPageFacClosure.TabIndex = 0
        Me.tbPageFacClosure.Text = "Closure Events"
        '
        'btnAddClosure
        '
        Me.btnAddClosure.Location = New System.Drawing.Point(0, 0)
        Me.btnAddClosure.Name = "btnAddClosure"
        Me.btnAddClosure.Size = New System.Drawing.Size(104, 23)
        Me.btnAddClosure.TabIndex = 2
        Me.btnAddClosure.Text = "Add Closure"
        Me.btnAddClosure.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dgClosureFacilityDetails
        '
        Me.dgClosureFacilityDetails.Cursor = System.Windows.Forms.Cursors.Default
        Me.dgClosureFacilityDetails.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.dgClosureFacilityDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgClosureFacilityDetails.Location = New System.Drawing.Point(0, 0)
        Me.dgClosureFacilityDetails.Name = "dgClosureFacilityDetails"
        Me.dgClosureFacilityDetails.Size = New System.Drawing.Size(946, 318)
        Me.dgClosureFacilityDetails.TabIndex = 1
        Me.dgClosureFacilityDetails.Text = "Closure Events"
        '
        'pnlFacilityClosureButton
        '
        Me.pnlFacilityClosureButton.Controls.Add(Me.lblNoOfClosuresValue)
        Me.pnlFacilityClosureButton.Controls.Add(Me.lblNoOfClosure)
        Me.pnlFacilityClosureButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFacilityClosureButton.Location = New System.Drawing.Point(0, 318)
        Me.pnlFacilityClosureButton.Name = "pnlFacilityClosureButton"
        Me.pnlFacilityClosureButton.Size = New System.Drawing.Size(946, 24)
        Me.pnlFacilityClosureButton.TabIndex = 98
        '
        'lblNoOfClosuresValue
        '
        Me.lblNoOfClosuresValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfClosuresValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfClosuresValue.Location = New System.Drawing.Point(208, 0)
        Me.lblNoOfClosuresValue.Name = "lblNoOfClosuresValue"
        Me.lblNoOfClosuresValue.Size = New System.Drawing.Size(48, 24)
        Me.lblNoOfClosuresValue.TabIndex = 5
        Me.lblNoOfClosuresValue.Text = "0"
        '
        'lblNoOfClosure
        '
        Me.lblNoOfClosure.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfClosure.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfClosure.Location = New System.Drawing.Point(0, 0)
        Me.lblNoOfClosure.Name = "lblNoOfClosure"
        Me.lblNoOfClosure.Size = New System.Drawing.Size(208, 24)
        Me.lblNoOfClosure.TabIndex = 4
        Me.lblNoOfClosure.Text = "Number of Closures at this Location:"
        '
        'tbPageFacilityDocuments
        '
        Me.tbPageFacilityDocuments.Controls.Add(Me.UCFacilityDocuments)
        Me.tbPageFacilityDocuments.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilityDocuments.Name = "tbPageFacilityDocuments"
        Me.tbPageFacilityDocuments.Size = New System.Drawing.Size(950, 354)
        Me.tbPageFacilityDocuments.TabIndex = 1
        Me.tbPageFacilityDocuments.Text = "Documents"
        '
        'UCFacilityDocuments
        '
        Me.UCFacilityDocuments.AutoScroll = True
        Me.UCFacilityDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCFacilityDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCFacilityDocuments.Name = "UCFacilityDocuments"
        Me.UCFacilityDocuments.Size = New System.Drawing.Size(950, 354)
        Me.UCFacilityDocuments.TabIndex = 2
        '
        'pnl_FacilityDetail
        '
        Me.pnl_FacilityDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnl_FacilityDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnl_FacilityDetail.Controls.Add(Me.lblInViolationValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.Label5)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFeeBalance)
        Me.pnl_FacilityDetail.Controls.Add(Me.LblFeeBalance)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtMGPTF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblMGPTF)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickAssess)
        Me.pnl_FacilityDetail.Controls.Add(Me.Label3)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnClosureFacLabels)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnClosureFacEnvelopes)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickUpcomingInstallDateValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblUpcomingInstallDate)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkUpcomingInstall)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblNextFac)
        Me.pnl_FacilityDetail.Controls.Add(Me.Panel2)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPStatusValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPStatus)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFuelBrand)
        Me.pnl_FacilityDetail.Controls.Add(Me.ll)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtFacilityPowerOff)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblPrevFacility)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateTransfered)
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
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickFacilityRecvd)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateReceived)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFuelBrandcmb)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacilityChangeCancel)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtDueByNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkSignatureofNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblPotentialOwner)
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
        Me.pnl_FacilityDetail.Controls.Add(Me.lblOwnerLastEditedBy)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblOwnerLastEditedOn)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityZip)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilitySIC)
        Me.pnl_FacilityDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_FacilityDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnl_FacilityDetail.Name = "pnl_FacilityDetail"
        Me.pnl_FacilityDetail.Size = New System.Drawing.Size(960, 288)
        Me.pnl_FacilityDetail.TabIndex = 98
        '
        'txtFeeBalance
        '
        Me.txtFeeBalance.AcceptsTab = True
        Me.txtFeeBalance.AutoSize = False
        Me.txtFeeBalance.Location = New System.Drawing.Point(696, 256)
        Me.txtFeeBalance.Name = "txtFeeBalance"
        Me.txtFeeBalance.ReadOnly = True
        Me.txtFeeBalance.Size = New System.Drawing.Size(120, 20)
        Me.txtFeeBalance.TabIndex = 1076
        Me.txtFeeBalance.Text = ""
        Me.txtFeeBalance.WordWrap = False
        '
        'LblFeeBalance
        '
        Me.LblFeeBalance.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblFeeBalance.Location = New System.Drawing.Point(616, 256)
        Me.LblFeeBalance.Name = "LblFeeBalance"
        Me.LblFeeBalance.Size = New System.Drawing.Size(80, 20)
        Me.LblFeeBalance.TabIndex = 1077
        Me.LblFeeBalance.Text = "Fee Balance"
        '
        'txtMGPTF
        '
        Me.txtMGPTF.AcceptsTab = True
        Me.txtMGPTF.AutoSize = False
        Me.txtMGPTF.Location = New System.Drawing.Point(696, 224)
        Me.txtMGPTF.Name = "txtMGPTF"
        Me.txtMGPTF.ReadOnly = True
        Me.txtMGPTF.Size = New System.Drawing.Size(120, 20)
        Me.txtMGPTF.TabIndex = 1074
        Me.txtMGPTF.Text = ""
        Me.txtMGPTF.WordWrap = False
        '
        'lblMGPTF
        '
        Me.lblMGPTF.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMGPTF.Location = New System.Drawing.Point(616, 224)
        Me.lblMGPTF.Name = "lblMGPTF"
        Me.lblMGPTF.Size = New System.Drawing.Size(80, 20)
        Me.lblMGPTF.TabIndex = 1075
        Me.lblMGPTF.Text = "MGPTF"
        '
        'dtPickAssess
        '
        Me.dtPickAssess.Checked = False
        Me.dtPickAssess.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickAssess.Location = New System.Drawing.Point(744, 200)
        Me.dtPickAssess.Name = "dtPickAssess"
        Me.dtPickAssess.ShowCheckBox = True
        Me.dtPickAssess.Size = New System.Drawing.Size(104, 21)
        Me.dtPickAssess.TabIndex = 1073
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(616, 200)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(152, 20)
        Me.Label3.TabIndex = 1072
        Me.Label3.Text = "TOS Assessment Date:"
        '
        'btnClosureFacLabels
        '
        Me.btnClosureFacLabels.Location = New System.Drawing.Point(11, 120)
        Me.btnClosureFacLabels.Name = "btnClosureFacLabels"
        Me.btnClosureFacLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnClosureFacLabels.TabIndex = 1070
        Me.btnClosureFacLabels.Text = "Labels"
        '
        'btnClosureFacEnvelopes
        '
        Me.btnClosureFacEnvelopes.Location = New System.Drawing.Point(11, 88)
        Me.btnClosureFacEnvelopes.Name = "btnClosureFacEnvelopes"
        Me.btnClosureFacEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnClosureFacEnvelopes.TabIndex = 1069
        Me.btnClosureFacEnvelopes.Text = "Envelopes"
        '
        'dtPickUpcomingInstallDateValue
        '
        Me.dtPickUpcomingInstallDateValue.Checked = False
        Me.dtPickUpcomingInstallDateValue.Enabled = False
        Me.dtPickUpcomingInstallDateValue.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickUpcomingInstallDateValue.Location = New System.Drawing.Point(504, 224)
        Me.dtPickUpcomingInstallDateValue.Name = "dtPickUpcomingInstallDateValue"
        Me.dtPickUpcomingInstallDateValue.ShowCheckBox = True
        Me.dtPickUpcomingInstallDateValue.Size = New System.Drawing.Size(101, 21)
        Me.dtPickUpcomingInstallDateValue.TabIndex = 1045
        '
        'lblUpcomingInstallDate
        '
        Me.lblUpcomingInstallDate.Location = New System.Drawing.Point(344, 224)
        Me.lblUpcomingInstallDate.Name = "lblUpcomingInstallDate"
        Me.lblUpcomingInstallDate.Size = New System.Drawing.Size(160, 20)
        Me.lblUpcomingInstallDate.TabIndex = 1044
        Me.lblUpcomingInstallDate.Text = "Upcoming Installation Date:"
        Me.lblUpcomingInstallDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkUpcomingInstall
        '
        Me.chkUpcomingInstall.Location = New System.Drawing.Point(192, 224)
        Me.chkUpcomingInstall.Name = "chkUpcomingInstall"
        Me.chkUpcomingInstall.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkUpcomingInstall.Size = New System.Drawing.Size(144, 20)
        Me.chkUpcomingInstall.TabIndex = 1043
        Me.chkUpcomingInstall.Text = "Upcoming Installation "
        Me.chkUpcomingInstall.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lnkLblNextFac
        '
        Me.lnkLblNextFac.Location = New System.Drawing.Point(272, 8)
        Me.lnkLblNextFac.Name = "lnkLblNextFac"
        Me.lnkLblNextFac.Size = New System.Drawing.Size(56, 16)
        Me.lnkLblNextFac.TabIndex = 1042
        Me.lnkLblNextFac.TabStop = True
        Me.lnkLblNextFac.Text = "Next>>"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnFacilityCancel)
        Me.Panel2.Controls.Add(Me.btnFacComments)
        Me.Panel2.Controls.Add(Me.btnFacilitySave)
        Me.Panel2.Controls.Add(Me.btnFacFlags)
        Me.Panel2.Location = New System.Drawing.Point(8, 248)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(352, 32)
        Me.Panel2.TabIndex = 1041
        '
        'btnFacilityCancel
        '
        Me.btnFacilityCancel.Enabled = False
        Me.btnFacilityCancel.Location = New System.Drawing.Point(108, 5)
        Me.btnFacilityCancel.Name = "btnFacilityCancel"
        Me.btnFacilityCancel.Size = New System.Drawing.Size(75, 26)
        Me.btnFacilityCancel.TabIndex = 30
        Me.btnFacilityCancel.Text = "Cancel"
        '
        'btnFacComments
        '
        Me.btnFacComments.Location = New System.Drawing.Point(268, 5)
        Me.btnFacComments.Name = "btnFacComments"
        Me.btnFacComments.Size = New System.Drawing.Size(80, 26)
        Me.btnFacComments.TabIndex = 1039
        Me.btnFacComments.Text = "Comments"
        '
        'btnFacilitySave
        '
        Me.btnFacilitySave.Enabled = False
        Me.btnFacilitySave.Location = New System.Drawing.Point(8, 4)
        Me.btnFacilitySave.Name = "btnFacilitySave"
        Me.btnFacilitySave.Size = New System.Drawing.Size(96, 26)
        Me.btnFacilitySave.TabIndex = 29
        Me.btnFacilitySave.Text = "Save Facility"
        Me.btnFacilitySave.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnFacFlags
        '
        Me.btnFacFlags.Location = New System.Drawing.Point(188, 5)
        Me.btnFacFlags.Name = "btnFacFlags"
        Me.btnFacFlags.Size = New System.Drawing.Size(75, 26)
        Me.btnFacFlags.TabIndex = 1040
        Me.btnFacFlags.Text = "Flags"
        '
        'lblCAPStatusValue
        '
        Me.lblCAPStatusValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblCAPStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCAPStatusValue.Enabled = False
        Me.lblCAPStatusValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCAPStatusValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblCAPStatusValue.Location = New System.Drawing.Point(456, 104)
        Me.lblCAPStatusValue.Name = "lblCAPStatusValue"
        Me.lblCAPStatusValue.Size = New System.Drawing.Size(120, 20)
        Me.lblCAPStatusValue.TabIndex = 1038
        '
        'lblCAPStatus
        '
        Me.lblCAPStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCAPStatus.Location = New System.Drawing.Point(344, 104)
        Me.lblCAPStatus.Name = "lblCAPStatus"
        Me.lblCAPStatus.Size = New System.Drawing.Size(100, 16)
        Me.lblCAPStatus.TabIndex = 1037
        Me.lblCAPStatus.Text = "CAP Status:"
        '
        'txtFuelBrand
        '
        Me.txtFuelBrand.Location = New System.Drawing.Point(456, 176)
        Me.txtFuelBrand.Name = "txtFuelBrand"
        Me.txtFuelBrand.Size = New System.Drawing.Size(72, 21)
        Me.txtFuelBrand.TabIndex = 1036
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
        Me.dtFacilityPowerOff.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtFacilityPowerOff.Location = New System.Drawing.Point(912, 248)
        Me.dtFacilityPowerOff.Name = "dtFacilityPowerOff"
        Me.dtFacilityPowerOff.ShowCheckBox = True
        Me.dtFacilityPowerOff.Size = New System.Drawing.Size(104, 21)
        Me.dtFacilityPowerOff.TabIndex = 1033
        Me.dtFacilityPowerOff.Visible = False
        '
        'lnkLblPrevFacility
        '
        Me.lnkLblPrevFacility.AutoSize = True
        Me.lnkLblPrevFacility.Location = New System.Drawing.Point(200, 8)
        Me.lnkLblPrevFacility.Name = "lnkLblPrevFacility"
        Me.lnkLblPrevFacility.Size = New System.Drawing.Size(70, 17)
        Me.lnkLblPrevFacility.TabIndex = 1031
        Me.lnkLblPrevFacility.TabStop = True
        Me.lnkLblPrevFacility.Text = "<< Previous"
        '
        'lblDateTransfered
        '
        Me.lblDateTransfered.Location = New System.Drawing.Point(896, 176)
        Me.lblDateTransfered.Name = "lblDateTransfered"
        Me.lblDateTransfered.Size = New System.Drawing.Size(32, 16)
        Me.lblDateTransfered.TabIndex = 1034
        Me.lblDateTransfered.Visible = False
        '
        'lblLUSTSite
        '
        Me.lblLUSTSite.Location = New System.Drawing.Point(344, 80)
        Me.lblLUSTSite.Name = "lblLUSTSite"
        Me.lblLUSTSite.Size = New System.Drawing.Size(104, 20)
        Me.lblLUSTSite.TabIndex = 1030
        Me.lblLUSTSite.Text = "Active LUST Site:"
        Me.lblLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkLUSTSite
        '
        Me.chkLUSTSite.Location = New System.Drawing.Point(456, 80)
        Me.chkLUSTSite.Name = "chkLUSTSite"
        Me.chkLUSTSite.Size = New System.Drawing.Size(16, 16)
        Me.chkLUSTSite.TabIndex = 28
        Me.chkLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPowerOff
        '
        Me.lblPowerOff.Location = New System.Drawing.Point(840, 248)
        Me.lblPowerOff.Name = "lblPowerOff"
        Me.lblPowerOff.Size = New System.Drawing.Size(80, 16)
        Me.lblPowerOff.TabIndex = 1028
        Me.lblPowerOff.Text = "Power Off:"
        Me.lblPowerOff.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblPowerOff.Visible = False
        '
        'lblCAPCandidate
        '
        Me.lblCAPCandidate.Location = New System.Drawing.Point(344, 128)
        Me.lblCAPCandidate.Name = "lblCAPCandidate"
        Me.lblCAPCandidate.Size = New System.Drawing.Size(100, 16)
        Me.lblCAPCandidate.TabIndex = 1026
        Me.lblCAPCandidate.Text = "CAP Candidate:"
        Me.lblCAPCandidate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkCAPCandidate
        '
        Me.chkCAPCandidate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCAPCandidate.Location = New System.Drawing.Point(456, 128)
        Me.chkCAPCandidate.Name = "chkCAPCandidate"
        Me.chkCAPCandidate.Size = New System.Drawing.Size(16, 16)
        Me.chkCAPCandidate.TabIndex = 26
        Me.chkCAPCandidate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFacilityLocationType
        '
        Me.lblFacilityLocationType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLocationType.Location = New System.Drawing.Point(616, 176)
        Me.lblFacilityLocationType.Name = "lblFacilityLocationType"
        Me.lblFacilityLocationType.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityLocationType.TabIndex = 1024
        Me.lblFacilityLocationType.Text = "Type:"
        '
        'cmbFacilityLocationType
        '
        Me.cmbFacilityLocationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityLocationType.DropDownWidth = 250
        Me.cmbFacilityLocationType.ItemHeight = 15
        Me.cmbFacilityLocationType.Location = New System.Drawing.Point(704, 176)
        Me.cmbFacilityLocationType.Name = "cmbFacilityLocationType"
        Me.cmbFacilityLocationType.Size = New System.Drawing.Size(144, 23)
        Me.cmbFacilityLocationType.TabIndex = 21
        '
        'lblFacilityMethod
        '
        Me.lblFacilityMethod.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityMethod.Location = New System.Drawing.Point(616, 152)
        Me.lblFacilityMethod.Name = "lblFacilityMethod"
        Me.lblFacilityMethod.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityMethod.TabIndex = 1022
        Me.lblFacilityMethod.Text = "Method:"
        '
        'cmbFacilityMethod
        '
        Me.cmbFacilityMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityMethod.DropDownWidth = 350
        Me.cmbFacilityMethod.ItemHeight = 15
        Me.cmbFacilityMethod.Location = New System.Drawing.Point(704, 152)
        Me.cmbFacilityMethod.Name = "cmbFacilityMethod"
        Me.cmbFacilityMethod.Size = New System.Drawing.Size(144, 23)
        Me.cmbFacilityMethod.TabIndex = 20
        '
        'lblFacilityDatum
        '
        Me.lblFacilityDatum.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityDatum.Location = New System.Drawing.Point(616, 128)
        Me.lblFacilityDatum.Name = "lblFacilityDatum"
        Me.lblFacilityDatum.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityDatum.TabIndex = 1020
        Me.lblFacilityDatum.Text = "Datum:"
        '
        'cmbFacilityDatum
        '
        Me.cmbFacilityDatum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityDatum.DropDownWidth = 250
        Me.cmbFacilityDatum.ItemHeight = 15
        Me.cmbFacilityDatum.Location = New System.Drawing.Point(704, 128)
        Me.cmbFacilityDatum.Name = "cmbFacilityDatum"
        Me.cmbFacilityDatum.Size = New System.Drawing.Size(144, 23)
        Me.cmbFacilityDatum.TabIndex = 19
        '
        'cmbFacilityType
        '
        Me.cmbFacilityType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityType.DropDownWidth = 180
        Me.cmbFacilityType.Enabled = False
        Me.cmbFacilityType.ItemHeight = 15
        Me.cmbFacilityType.Location = New System.Drawing.Point(704, 56)
        Me.cmbFacilityType.Name = "cmbFacilityType"
        Me.cmbFacilityType.Size = New System.Drawing.Size(144, 23)
        Me.cmbFacilityType.TabIndex = 12
        '
        'txtFacilityLatSec
        '
        Me.txtFacilityLatSec.AcceptsTab = True
        Me.txtFacilityLatSec.AutoSize = False
        Me.txtFacilityLatSec.Enabled = False
        Me.txtFacilityLatSec.Location = New System.Drawing.Point(796, 80)
        Me.txtFacilityLatSec.MaxLength = 5
        Me.txtFacilityLatSec.Name = "txtFacilityLatSec"
        Me.txtFacilityLatSec.Size = New System.Drawing.Size(37, 21)
        Me.txtFacilityLatSec.TabIndex = 15
        Me.txtFacilityLatSec.Text = ""
        Me.txtFacilityLatSec.WordWrap = False
        '
        'txtFacilityLongSec
        '
        Me.txtFacilityLongSec.AcceptsTab = True
        Me.txtFacilityLongSec.AutoSize = False
        Me.txtFacilityLongSec.Enabled = False
        Me.txtFacilityLongSec.Location = New System.Drawing.Point(796, 104)
        Me.txtFacilityLongSec.MaxLength = 5
        Me.txtFacilityLongSec.Name = "txtFacilityLongSec"
        Me.txtFacilityLongSec.Size = New System.Drawing.Size(38, 21)
        Me.txtFacilityLongSec.TabIndex = 18
        Me.txtFacilityLongSec.Text = ""
        Me.txtFacilityLongSec.WordWrap = False
        '
        'txtFacilityLatMin
        '
        Me.txtFacilityLatMin.AcceptsTab = True
        Me.txtFacilityLatMin.AutoSize = False
        Me.txtFacilityLatMin.Enabled = False
        Me.txtFacilityLatMin.Location = New System.Drawing.Point(752, 80)
        Me.txtFacilityLatMin.MaxLength = 2
        Me.txtFacilityLatMin.Name = "txtFacilityLatMin"
        Me.txtFacilityLatMin.Size = New System.Drawing.Size(28, 21)
        Me.txtFacilityLatMin.TabIndex = 14
        Me.txtFacilityLatMin.Text = ""
        Me.txtFacilityLatMin.WordWrap = False
        '
        'txtFacilityLongMin
        '
        Me.txtFacilityLongMin.AcceptsTab = True
        Me.txtFacilityLongMin.AutoSize = False
        Me.txtFacilityLongMin.Enabled = False
        Me.txtFacilityLongMin.Location = New System.Drawing.Point(752, 104)
        Me.txtFacilityLongMin.MaxLength = 2
        Me.txtFacilityLongMin.Name = "txtFacilityLongMin"
        Me.txtFacilityLongMin.Size = New System.Drawing.Size(28, 21)
        Me.txtFacilityLongMin.TabIndex = 17
        Me.txtFacilityLongMin.Text = ""
        Me.txtFacilityLongMin.WordWrap = False
        '
        'lblFacilityLongMin
        '
        Me.lblFacilityLongMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongMin.Location = New System.Drawing.Point(784, 80)
        Me.lblFacilityLongMin.Name = "lblFacilityLongMin"
        Me.lblFacilityLongMin.Size = New System.Drawing.Size(8, 20)
        Me.lblFacilityLongMin.TabIndex = 1018
        Me.lblFacilityLongMin.Text = "'"
        '
        'lblFacilityLongSec
        '
        Me.lblFacilityLongSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongSec.Location = New System.Drawing.Point(836, 104)
        Me.lblFacilityLongSec.Name = "lblFacilityLongSec"
        Me.lblFacilityLongSec.Size = New System.Drawing.Size(10, 20)
        Me.lblFacilityLongSec.TabIndex = 1017
        Me.lblFacilityLongSec.Text = """"
        '
        'lblFacilityLatMin
        '
        Me.lblFacilityLatMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatMin.Location = New System.Drawing.Point(784, 104)
        Me.lblFacilityLatMin.Name = "lblFacilityLatMin"
        Me.lblFacilityLatMin.Size = New System.Drawing.Size(8, 20)
        Me.lblFacilityLatMin.TabIndex = 1016
        Me.lblFacilityLatMin.Text = "'"
        '
        'lblFacilityLatSec
        '
        Me.lblFacilityLatSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatSec.Location = New System.Drawing.Point(836, 80)
        Me.lblFacilityLatSec.Name = "lblFacilityLatSec"
        Me.lblFacilityLatSec.Size = New System.Drawing.Size(10, 20)
        Me.lblFacilityLatSec.TabIndex = 1015
        Me.lblFacilityLatSec.Text = """"
        '
        'lblFacilityLongDegree
        '
        Me.lblFacilityLongDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongDegree.Location = New System.Drawing.Point(736, 96)
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
        Me.mskTxtFacilityFax.Size = New System.Drawing.Size(120, 23)
        Me.mskTxtFacilityFax.TabIndex = 11
        '
        'mskTxtFacilityPhone
        '
        Me.mskTxtFacilityPhone.ContainingControl = Me
        Me.mskTxtFacilityPhone.Location = New System.Drawing.Point(88, 162)
        Me.mskTxtFacilityPhone.Name = "mskTxtFacilityPhone"
        Me.mskTxtFacilityPhone.OcxState = CType(resources.GetObject("mskTxtFacilityPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFacilityPhone.Size = New System.Drawing.Size(120, 23)
        Me.mskTxtFacilityPhone.TabIndex = 10
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
        Me.lblFacilitySIC.Location = New System.Drawing.Point(344, 152)
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
        Me.lblFacilityFax.Size = New System.Drawing.Size(72, 20)
        Me.lblFacilityFax.TabIndex = 147
        Me.lblFacilityFax.Text = "Fax:"
        '
        'dtPickFacilityRecvd
        '
        Me.dtPickFacilityRecvd.Checked = False
        Me.dtPickFacilityRecvd.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickFacilityRecvd.Location = New System.Drawing.Point(456, 32)
        Me.dtPickFacilityRecvd.Name = "dtPickFacilityRecvd"
        Me.dtPickFacilityRecvd.ShowCheckBox = True
        Me.dtPickFacilityRecvd.Size = New System.Drawing.Size(104, 21)
        Me.dtPickFacilityRecvd.TabIndex = 22
        '
        'lblDateReceived
        '
        Me.lblDateReceived.Location = New System.Drawing.Point(344, 32)
        Me.lblDateReceived.Name = "lblDateReceived"
        Me.lblDateReceived.Size = New System.Drawing.Size(100, 20)
        Me.lblDateReceived.TabIndex = 145
        Me.lblDateReceived.Text = "Date Received:"
        '
        'txtFuelBrandcmb
        '
        Me.txtFuelBrandcmb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple
        Me.txtFuelBrandcmb.ItemHeight = 15
        Me.txtFuelBrandcmb.Items.AddRange(New Object() {"Shell", "Exon", "Texaco", "Mobil"})
        Me.txtFuelBrandcmb.Location = New System.Drawing.Point(872, 88)
        Me.txtFuelBrandcmb.Name = "txtFuelBrandcmb"
        Me.txtFuelBrandcmb.Size = New System.Drawing.Size(96, 23)
        Me.txtFuelBrandcmb.TabIndex = 23
        Me.txtFuelBrandcmb.Text = "Remove after cure period"
        Me.txtFuelBrandcmb.Visible = False
        '
        'btnFacilityChangeCancel
        '
        Me.btnFacilityChangeCancel.Enabled = False
        Me.btnFacilityChangeCancel.Location = New System.Drawing.Point(880, 200)
        Me.btnFacilityChangeCancel.Name = "btnFacilityChangeCancel"
        Me.btnFacilityChangeCancel.TabIndex = 30
        Me.btnFacilityChangeCancel.Text = "Cancel"
        Me.btnFacilityChangeCancel.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnFacilityChangeCancel.Visible = False
        '
        'txtDueByNF
        '
        Me.txtDueByNF.AcceptsTab = True
        Me.txtDueByNF.AutoSize = False
        Me.txtDueByNF.Enabled = False
        Me.txtDueByNF.Location = New System.Drawing.Point(936, 176)
        Me.txtDueByNF.Name = "txtDueByNF"
        Me.txtDueByNF.Size = New System.Drawing.Size(64, 21)
        Me.txtDueByNF.TabIndex = 136
        Me.txtDueByNF.Text = ""
        Me.txtDueByNF.Visible = False
        Me.txtDueByNF.WordWrap = False
        '
        'chkSignatureofNF
        '
        Me.chkSignatureofNF.Location = New System.Drawing.Point(128, 224)
        Me.chkSignatureofNF.Name = "chkSignatureofNF"
        Me.chkSignatureofNF.Size = New System.Drawing.Size(16, 16)
        Me.chkSignatureofNF.TabIndex = 24
        Me.chkSignatureofNF.Text = "CheckBox5"
        '
        'lblPotentialOwner
        '
        Me.lblPotentialOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPotentialOwner.ForeColor = System.Drawing.Color.Red
        Me.lblPotentialOwner.Location = New System.Drawing.Point(912, 208)
        Me.lblPotentialOwner.Name = "lblPotentialOwner"
        Me.lblPotentialOwner.Size = New System.Drawing.Size(128, 23)
        Me.lblPotentialOwner.TabIndex = 129
        Me.lblPotentialOwner.Text = "Potential Owner:"
        Me.lblPotentialOwner.Visible = False
        '
        'lblFacilitySigOnNF
        '
        Me.lblFacilitySigOnNF.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySigOnNF.Location = New System.Drawing.Point(8, 224)
        Me.lblFacilitySigOnNF.Name = "lblFacilitySigOnNF"
        Me.lblFacilitySigOnNF.Size = New System.Drawing.Size(120, 20)
        Me.lblFacilitySigOnNF.TabIndex = 127
        Me.lblFacilitySigOnNF.Text = "Signature Received:"
        '
        'lblFacilityFuelBrand
        '
        Me.lblFacilityFuelBrand.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityFuelBrand.Location = New System.Drawing.Point(344, 176)
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
        Me.lblFacilityStatusValue.Location = New System.Drawing.Point(456, 56)
        Me.lblFacilityStatusValue.Name = "lblFacilityStatusValue"
        Me.lblFacilityStatusValue.Size = New System.Drawing.Size(120, 20)
        Me.lblFacilityStatusValue.TabIndex = 124
        '
        'lblFacilityStatus
        '
        Me.lblFacilityStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityStatus.Location = New System.Drawing.Point(344, 56)
        Me.lblFacilityStatus.Name = "lblFacilityStatus"
        Me.lblFacilityStatus.Size = New System.Drawing.Size(100, 20)
        Me.lblFacilityStatus.TabIndex = 123
        Me.lblFacilityStatus.Text = "Facility Status:"
        '
        'txtFacilityLongDegree
        '
        Me.txtFacilityLongDegree.AcceptsTab = True
        Me.txtFacilityLongDegree.AutoSize = False
        Me.txtFacilityLongDegree.Enabled = False
        Me.txtFacilityLongDegree.Location = New System.Drawing.Point(704, 104)
        Me.txtFacilityLongDegree.MaxLength = 3
        Me.txtFacilityLongDegree.Name = "txtFacilityLongDegree"
        Me.txtFacilityLongDegree.Size = New System.Drawing.Size(32, 21)
        Me.txtFacilityLongDegree.TabIndex = 16
        Me.txtFacilityLongDegree.Text = ""
        Me.txtFacilityLongDegree.WordWrap = False
        '
        'txtFacilityLatDegree
        '
        Me.txtFacilityLatDegree.AcceptsTab = True
        Me.txtFacilityLatDegree.AutoSize = False
        Me.txtFacilityLatDegree.Enabled = False
        Me.txtFacilityLatDegree.Location = New System.Drawing.Point(704, 80)
        Me.txtFacilityLatDegree.MaxLength = 3
        Me.txtFacilityLatDegree.Name = "txtFacilityLatDegree"
        Me.txtFacilityLatDegree.Size = New System.Drawing.Size(32, 21)
        Me.txtFacilityLatDegree.TabIndex = 13
        Me.txtFacilityLatDegree.Text = ""
        Me.txtFacilityLatDegree.WordWrap = False
        '
        'lblFacilityLongitude
        '
        Me.lblFacilityLongitude.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongitude.Location = New System.Drawing.Point(616, 104)
        Me.lblFacilityLongitude.Name = "lblFacilityLongitude"
        Me.lblFacilityLongitude.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityLongitude.TabIndex = 111
        Me.lblFacilityLongitude.Text = "Longitude:"
        '
        'lblFacilityLatitude
        '
        Me.lblFacilityLatitude.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatitude.Location = New System.Drawing.Point(616, 80)
        Me.lblFacilityLatitude.Name = "lblFacilityLatitude"
        Me.lblFacilityLatitude.Size = New System.Drawing.Size(80, 20)
        Me.lblFacilityLatitude.TabIndex = 110
        Me.lblFacilityLatitude.Text = "Latitude:"
        '
        'lblFacilityType
        '
        Me.lblFacilityType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityType.Location = New System.Drawing.Point(616, 56)
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
        Me.txtFacilityAIID.Location = New System.Drawing.Point(704, 32)
        Me.txtFacilityAIID.Name = "txtFacilityAIID"
        Me.txtFacilityAIID.Size = New System.Drawing.Size(144, 20)
        Me.txtFacilityAIID.TabIndex = 105
        Me.txtFacilityAIID.Text = ""
        Me.txtFacilityAIID.WordWrap = False
        '
        'lblfacilityAIID
        '
        Me.lblfacilityAIID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblfacilityAIID.Location = New System.Drawing.Point(616, 32)
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
        Me.lblFacilityIDValue.Size = New System.Drawing.Size(88, 20)
        Me.lblFacilityIDValue.TabIndex = 103
        '
        'lblFacilityID
        '
        Me.lblFacilityID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityID.Location = New System.Drawing.Point(8, 8)
        Me.lblFacilityID.Name = "lblFacilityID"
        Me.lblFacilityID.Size = New System.Drawing.Size(70, 20)
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
        Me.lblFacilityPhone.Size = New System.Drawing.Size(72, 20)
        Me.lblFacilityPhone.TabIndex = 98
        Me.lblFacilityPhone.Text = "Phone:"
        Me.lblFacilityPhone.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFacilityName
        '
        Me.txtFacilityName.AcceptsTab = True
        Me.txtFacilityName.AutoSize = False
        Me.txtFacilityName.Enabled = False
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
        Me.lblFacilityAddress.Size = New System.Drawing.Size(72, 20)
        Me.lblFacilityAddress.TabIndex = 90
        Me.lblFacilityAddress.Text = "Address:"
        '
        'lblFacilityName
        '
        Me.lblFacilityName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityName.Location = New System.Drawing.Point(8, 32)
        Me.lblFacilityName.Name = "lblFacilityName"
        Me.lblFacilityName.Size = New System.Drawing.Size(90, 20)
        Me.lblFacilityName.TabIndex = 89
        Me.lblFacilityName.Text = "Facility Name:"
        Me.lblFacilityName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFacilityZip
        '
        Me.txtFacilityZip.Location = New System.Drawing.Point(896, 56)
        Me.txtFacilityZip.Name = "txtFacilityZip"
        Me.txtFacilityZip.Size = New System.Drawing.Size(96, 21)
        Me.txtFacilityZip.TabIndex = 11
        Me.txtFacilityZip.Text = ""
        Me.txtFacilityZip.Visible = False
        '
        'lblFacilityLatDegree
        '
        Me.lblFacilityLatDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatDegree.Location = New System.Drawing.Point(736, 72)
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
        Me.txtFacilitySIC.Location = New System.Drawing.Point(456, 152)
        Me.txtFacilitySIC.Name = "txtFacilitySIC"
        Me.txtFacilitySIC.Size = New System.Drawing.Size(120, 20)
        Me.txtFacilitySIC.TabIndex = 1038
        '
        'tbPageClosures
        '
        Me.tbPageClosures.Controls.Add(Me.pnlClosureMain)
        Me.tbPageClosures.Controls.Add(Me.pnlLUSTEventBottom)
        Me.tbPageClosures.Controls.Add(Me.pnlNoticeOfInterestHeader)
        Me.tbPageClosures.Location = New System.Drawing.Point(4, 22)
        Me.tbPageClosures.Name = "tbPageClosures"
        Me.tbPageClosures.Size = New System.Drawing.Size(964, 668)
        Me.tbPageClosures.TabIndex = 12
        Me.tbPageClosures.Text = "Closures"
        Me.tbPageClosures.Visible = False
        '
        'pnlClosureMain
        '
        Me.pnlClosureMain.AutoScroll = True
        Me.pnlClosureMain.Controls.Add(Me.pnlClosuresDetails)
        Me.pnlClosureMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlClosureMain.Location = New System.Drawing.Point(0, 64)
        Me.pnlClosureMain.Name = "pnlClosureMain"
        Me.pnlClosureMain.Size = New System.Drawing.Size(964, 564)
        Me.pnlClosureMain.TabIndex = 3
        '
        'pnlClosuresDetails
        '
        Me.pnlClosuresDetails.AutoScroll = True
        Me.pnlClosuresDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlClosuresDetails.Controls.Add(Me.pnlContactDetails)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlContacts)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlClosureReportDetails)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlClosureReport)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlChecklistDetails)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlChecklist)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlAnalysisDetails)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlAnalysis)
        Me.pnlClosuresDetails.Controls.Add(Me.PnlTanksPipesDetails)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlTanksPipes)
        Me.pnlClosuresDetails.Controls.Add(Me.PnlNoticeOfInterestDetails)
        Me.pnlClosuresDetails.Controls.Add(Me.pnlNoticeOfInterest)
        Me.pnlClosuresDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlClosuresDetails.DockPadding.All = 3
        Me.pnlClosuresDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlClosuresDetails.Name = "pnlClosuresDetails"
        Me.pnlClosuresDetails.Size = New System.Drawing.Size(964, 564)
        Me.pnlClosuresDetails.TabIndex = 2
        '
        'pnlContactDetails
        '
        Me.pnlContactDetails.Controls.Add(Me.pnlClosureContactButtons)
        Me.pnlContactDetails.Controls.Add(Me.pnlClosureContactsContainer)
        Me.pnlContactDetails.Controls.Add(Me.pnlClosureContactHeader)
        Me.pnlContactDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContactDetails.Location = New System.Drawing.Point(3, 1256)
        Me.pnlContactDetails.Name = "pnlContactDetails"
        Me.pnlContactDetails.Size = New System.Drawing.Size(954, 312)
        Me.pnlContactDetails.TabIndex = 201
        '
        'pnlClosureContactButtons
        '
        Me.pnlClosureContactButtons.Controls.Add(Me.btnClosureContactModify)
        Me.pnlClosureContactButtons.Controls.Add(Me.btnClosureContactDelete)
        Me.pnlClosureContactButtons.Controls.Add(Me.btnClosureContactAssociate)
        Me.pnlClosureContactButtons.Controls.Add(Me.btnClosureContactAddorSearch)
        Me.pnlClosureContactButtons.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlClosureContactButtons.DockPadding.All = 3
        Me.pnlClosureContactButtons.Location = New System.Drawing.Point(0, 256)
        Me.pnlClosureContactButtons.Name = "pnlClosureContactButtons"
        Me.pnlClosureContactButtons.Size = New System.Drawing.Size(954, 48)
        Me.pnlClosureContactButtons.TabIndex = 2
        '
        'btnClosureContactModify
        '
        Me.btnClosureContactModify.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClosureContactModify.Location = New System.Drawing.Point(240, 13)
        Me.btnClosureContactModify.Name = "btnClosureContactModify"
        Me.btnClosureContactModify.Size = New System.Drawing.Size(235, 26)
        Me.btnClosureContactModify.TabIndex = 1
        Me.btnClosureContactModify.Text = "Modify Contact"
        '
        'btnClosureContactDelete
        '
        Me.btnClosureContactDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClosureContactDelete.Location = New System.Drawing.Point(472, 13)
        Me.btnClosureContactDelete.Name = "btnClosureContactDelete"
        Me.btnClosureContactDelete.Size = New System.Drawing.Size(235, 26)
        Me.btnClosureContactDelete.TabIndex = 2
        Me.btnClosureContactDelete.Text = "Disassociate Contact"
        '
        'btnClosureContactAssociate
        '
        Me.btnClosureContactAssociate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClosureContactAssociate.Location = New System.Drawing.Point(704, 13)
        Me.btnClosureContactAssociate.Name = "btnClosureContactAssociate"
        Me.btnClosureContactAssociate.Size = New System.Drawing.Size(235, 26)
        Me.btnClosureContactAssociate.TabIndex = 3
        Me.btnClosureContactAssociate.Text = "Associate Contact from Different Module"
        '
        'btnClosureContactAddorSearch
        '
        Me.btnClosureContactAddorSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClosureContactAddorSearch.Location = New System.Drawing.Point(8, 13)
        Me.btnClosureContactAddorSearch.Name = "btnClosureContactAddorSearch"
        Me.btnClosureContactAddorSearch.Size = New System.Drawing.Size(235, 26)
        Me.btnClosureContactAddorSearch.TabIndex = 0
        Me.btnClosureContactAddorSearch.Text = "Add/Search Contact to Associate"
        '
        'pnlClosureContactsContainer
        '
        Me.pnlClosureContactsContainer.AutoScroll = True
        Me.pnlClosureContactsContainer.Controls.Add(Me.ugClosureContacts)
        Me.pnlClosureContactsContainer.Controls.Add(Me.Label4)
        Me.pnlClosureContactsContainer.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlClosureContactsContainer.Location = New System.Drawing.Point(0, 32)
        Me.pnlClosureContactsContainer.Name = "pnlClosureContactsContainer"
        Me.pnlClosureContactsContainer.Size = New System.Drawing.Size(954, 224)
        Me.pnlClosureContactsContainer.TabIndex = 1
        '
        'ugClosureContacts
        '
        Me.ugClosureContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugClosureContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugClosureContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugClosureContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugClosureContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugClosureContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugClosureContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugClosureContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugClosureContacts.Name = "ugClosureContacts"
        Me.ugClosureContacts.Size = New System.Drawing.Size(954, 224)
        Me.ugClosureContacts.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(792, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(7, 23)
        Me.Label4.TabIndex = 2
        '
        'pnlClosureContactHeader
        '
        Me.pnlClosureContactHeader.Controls.Add(Me.chkClosureShowActive)
        Me.pnlClosureContactHeader.Controls.Add(Me.chkClosureShowRelatedContacts)
        Me.pnlClosureContactHeader.Controls.Add(Me.chkClosureShowContactsForAllModules)
        Me.pnlClosureContactHeader.Controls.Add(Me.lblClosureContacts)
        Me.pnlClosureContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlClosureContactHeader.DockPadding.All = 3
        Me.pnlClosureContactHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlClosureContactHeader.Name = "pnlClosureContactHeader"
        Me.pnlClosureContactHeader.Size = New System.Drawing.Size(954, 32)
        Me.pnlClosureContactHeader.TabIndex = 0
        '
        'chkClosureShowActive
        '
        Me.chkClosureShowActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkClosureShowActive.Location = New System.Drawing.Point(635, 8)
        Me.chkClosureShowActive.Name = "chkClosureShowActive"
        Me.chkClosureShowActive.Size = New System.Drawing.Size(144, 16)
        Me.chkClosureShowActive.TabIndex = 2
        Me.chkClosureShowActive.Tag = "646"
        Me.chkClosureShowActive.Text = "Show Active Only"
        '
        'chkClosureShowRelatedContacts
        '
        Me.chkClosureShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkClosureShowRelatedContacts.Location = New System.Drawing.Point(467, 8)
        Me.chkClosureShowRelatedContacts.Name = "chkClosureShowRelatedContacts"
        Me.chkClosureShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkClosureShowRelatedContacts.TabIndex = 1
        Me.chkClosureShowRelatedContacts.Tag = "645"
        Me.chkClosureShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkClosureShowContactsForAllModules
        '
        Me.chkClosureShowContactsForAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkClosureShowContactsForAllModules.Location = New System.Drawing.Point(251, 8)
        Me.chkClosureShowContactsForAllModules.Name = "chkClosureShowContactsForAllModules"
        Me.chkClosureShowContactsForAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkClosureShowContactsForAllModules.TabIndex = 0
        Me.chkClosureShowContactsForAllModules.Tag = "644"
        Me.chkClosureShowContactsForAllModules.Text = "Show Contacts for All Modules"
        '
        'lblClosureContacts
        '
        Me.lblClosureContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosureContacts.Location = New System.Drawing.Point(8, 8)
        Me.lblClosureContacts.Name = "lblClosureContacts"
        Me.lblClosureContacts.Size = New System.Drawing.Size(104, 16)
        Me.lblClosureContacts.TabIndex = 139
        Me.lblClosureContacts.Text = "Closure Contacts"
        '
        'pnlContacts
        '
        Me.pnlContacts.Controls.Add(Me.lblContactsHead)
        Me.pnlContacts.Controls.Add(Me.lblContactsDisplay)
        Me.pnlContacts.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContacts.Location = New System.Drawing.Point(3, 1232)
        Me.pnlContacts.Name = "pnlContacts"
        Me.pnlContacts.Size = New System.Drawing.Size(954, 24)
        Me.pnlContacts.TabIndex = 200
        '
        'lblContactsHead
        '
        Me.lblContactsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblContactsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblContactsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblContactsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblContactsHead.Name = "lblContactsHead"
        Me.lblContactsHead.Size = New System.Drawing.Size(938, 24)
        Me.lblContactsHead.TabIndex = 1
        Me.lblContactsHead.Text = "Contacts"
        Me.lblContactsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblContactsDisplay
        '
        Me.lblContactsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblContactsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblContactsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblContactsDisplay.Name = "lblContactsDisplay"
        Me.lblContactsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblContactsDisplay.TabIndex = 0
        Me.lblContactsDisplay.Text = "-"
        Me.lblContactsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlClosureReportDetails
        '
        Me.pnlClosureReportDetails.Controls.Add(Me.TxtInspections)
        Me.pnlClosureReportDetails.Controls.Add(Me.LblInspections)
        Me.pnlClosureReportDetails.Controls.Add(Me.lblLicenseeSearch)
        Me.pnlClosureReportDetails.Controls.Add(Me.txtLicensee)
        Me.pnlClosureReportDetails.Controls.Add(Me.txtClosureReportCompany)
        Me.pnlClosureReportDetails.Controls.Add(Me.dtPickClosureReceived)
        Me.pnlClosureReportDetails.Controls.Add(Me.btnProcessClosure)
        Me.pnlClosureReportDetails.Controls.Add(Me.lblClosureReceived)
        Me.pnlClosureReportDetails.Controls.Add(Me.cmbClosureReportFillMaterial)
        Me.pnlClosureReportDetails.Controls.Add(Me.lblClosureReportCertifiedContractor)
        Me.pnlClosureReportDetails.Controls.Add(Me.lblClosureReportDateClosed)
        Me.pnlClosureReportDetails.Controls.Add(Me.lblDateLastUsed)
        Me.pnlClosureReportDetails.Controls.Add(Me.lblClosureReportCompany)
        Me.pnlClosureReportDetails.Controls.Add(Me.lblClosureReportFillMaterial)
        Me.pnlClosureReportDetails.Controls.Add(Me.dtPickDateClosed)
        Me.pnlClosureReportDetails.Controls.Add(Me.dtPickDateLastUsed)
        Me.pnlClosureReportDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlClosureReportDetails.Location = New System.Drawing.Point(3, 1104)
        Me.pnlClosureReportDetails.Name = "pnlClosureReportDetails"
        Me.pnlClosureReportDetails.Size = New System.Drawing.Size(954, 128)
        Me.pnlClosureReportDetails.TabIndex = 58
        '
        'TxtInspections
        '
        Me.TxtInspections.Location = New System.Drawing.Point(664, 56)
        Me.TxtInspections.Multiline = True
        Me.TxtInspections.Name = "TxtInspections"
        Me.TxtInspections.ReadOnly = True
        Me.TxtInspections.Size = New System.Drawing.Size(264, 21)
        Me.TxtInspections.TabIndex = 224
        Me.TxtInspections.Text = ""
        '
        'LblInspections
        '
        Me.LblInspections.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblInspections.Location = New System.Drawing.Point(656, 40)
        Me.LblInspections.Name = "LblInspections"
        Me.LblInspections.Size = New System.Drawing.Size(136, 16)
        Me.LblInspections.TabIndex = 223
        Me.LblInspections.Text = "Inspection Scheduled:"
        Me.LblInspections.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblLicenseeSearch
        '
        Me.lblLicenseeSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLicenseeSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLicenseeSearch.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblLicenseeSearch.Location = New System.Drawing.Point(632, 16)
        Me.lblLicenseeSearch.Name = "lblLicenseeSearch"
        Me.lblLicenseeSearch.Size = New System.Drawing.Size(16, 22)
        Me.lblLicenseeSearch.TabIndex = 222
        Me.lblLicenseeSearch.Text = "?"
        Me.lblLicenseeSearch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtLicensee
        '
        Me.txtLicensee.Location = New System.Drawing.Point(424, 16)
        Me.txtLicensee.Name = "txtLicensee"
        Me.txtLicensee.Size = New System.Drawing.Size(200, 21)
        Me.txtLicensee.TabIndex = 187
        Me.txtLicensee.Text = ""
        '
        'txtClosureReportCompany
        '
        Me.txtClosureReportCompany.Location = New System.Drawing.Point(424, 48)
        Me.txtClosureReportCompany.Name = "txtClosureReportCompany"
        Me.txtClosureReportCompany.Size = New System.Drawing.Size(200, 21)
        Me.txtClosureReportCompany.TabIndex = 186
        Me.txtClosureReportCompany.Text = ""
        '
        'dtPickClosureReceived
        '
        Me.dtPickClosureReceived.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosureReceived.Checked = False
        Me.dtPickClosureReceived.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosureReceived.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosureReceived.Location = New System.Drawing.Point(152, 48)
        Me.dtPickClosureReceived.Name = "dtPickClosureReceived"
        Me.dtPickClosureReceived.ShowCheckBox = True
        Me.dtPickClosureReceived.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosureReceived.TabIndex = 0
        Me.dtPickClosureReceived.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'btnProcessClosure
        '
        Me.btnProcessClosure.Enabled = False
        Me.btnProcessClosure.Location = New System.Drawing.Point(664, 8)
        Me.btnProcessClosure.Name = "btnProcessClosure"
        Me.btnProcessClosure.Size = New System.Drawing.Size(208, 24)
        Me.btnProcessClosure.TabIndex = 6
        Me.btnProcessClosure.Text = "Process Closure"
        '
        'lblClosureReceived
        '
        Me.lblClosureReceived.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosureReceived.Location = New System.Drawing.Point(24, 48)
        Me.lblClosureReceived.Name = "lblClosureReceived"
        Me.lblClosureReceived.Size = New System.Drawing.Size(120, 16)
        Me.lblClosureReceived.TabIndex = 185
        Me.lblClosureReceived.Text = "Closure Received:"
        Me.lblClosureReceived.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'cmbClosureReportFillMaterial
        '
        Me.cmbClosureReportFillMaterial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbClosureReportFillMaterial.DropDownWidth = 152
        Me.cmbClosureReportFillMaterial.Enabled = False
        Me.cmbClosureReportFillMaterial.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbClosureReportFillMaterial.ItemHeight = 15
        Me.cmbClosureReportFillMaterial.Location = New System.Drawing.Point(424, 80)
        Me.cmbClosureReportFillMaterial.Name = "cmbClosureReportFillMaterial"
        Me.cmbClosureReportFillMaterial.Size = New System.Drawing.Size(152, 23)
        Me.cmbClosureReportFillMaterial.TabIndex = 5
        '
        'lblClosureReportCertifiedContractor
        '
        Me.lblClosureReportCertifiedContractor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosureReportCertifiedContractor.Location = New System.Drawing.Point(280, 16)
        Me.lblClosureReportCertifiedContractor.Name = "lblClosureReportCertifiedContractor"
        Me.lblClosureReportCertifiedContractor.Size = New System.Drawing.Size(136, 16)
        Me.lblClosureReportCertifiedContractor.TabIndex = 185
        Me.lblClosureReportCertifiedContractor.Text = "Certified Contractor:"
        Me.lblClosureReportCertifiedContractor.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblClosureReportDateClosed
        '
        Me.lblClosureReportDateClosed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosureReportDateClosed.Location = New System.Drawing.Point(24, 80)
        Me.lblClosureReportDateClosed.Name = "lblClosureReportDateClosed"
        Me.lblClosureReportDateClosed.Size = New System.Drawing.Size(120, 16)
        Me.lblClosureReportDateClosed.TabIndex = 185
        Me.lblClosureReportDateClosed.Text = "Date Closed:"
        Me.lblClosureReportDateClosed.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblDateLastUsed
        '
        Me.lblDateLastUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDateLastUsed.Location = New System.Drawing.Point(24, 16)
        Me.lblDateLastUsed.Name = "lblDateLastUsed"
        Me.lblDateLastUsed.Size = New System.Drawing.Size(120, 16)
        Me.lblDateLastUsed.TabIndex = 185
        Me.lblDateLastUsed.Text = "Date Late Used:"
        Me.lblDateLastUsed.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblClosureReportCompany
        '
        Me.lblClosureReportCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosureReportCompany.Location = New System.Drawing.Point(280, 48)
        Me.lblClosureReportCompany.Name = "lblClosureReportCompany"
        Me.lblClosureReportCompany.Size = New System.Drawing.Size(136, 16)
        Me.lblClosureReportCompany.TabIndex = 185
        Me.lblClosureReportCompany.Text = "Company:"
        Me.lblClosureReportCompany.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblClosureReportFillMaterial
        '
        Me.lblClosureReportFillMaterial.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosureReportFillMaterial.Location = New System.Drawing.Point(280, 80)
        Me.lblClosureReportFillMaterial.Name = "lblClosureReportFillMaterial"
        Me.lblClosureReportFillMaterial.Size = New System.Drawing.Size(136, 16)
        Me.lblClosureReportFillMaterial.TabIndex = 185
        Me.lblClosureReportFillMaterial.Text = "Fill Material:"
        Me.lblClosureReportFillMaterial.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'dtPickDateClosed
        '
        Me.dtPickDateClosed.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickDateClosed.Checked = False
        Me.dtPickDateClosed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickDateClosed.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickDateClosed.Location = New System.Drawing.Point(152, 80)
        Me.dtPickDateClosed.Name = "dtPickDateClosed"
        Me.dtPickDateClosed.ShowCheckBox = True
        Me.dtPickDateClosed.Size = New System.Drawing.Size(104, 21)
        Me.dtPickDateClosed.TabIndex = 1
        Me.dtPickDateClosed.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickDateLastUsed
        '
        Me.dtPickDateLastUsed.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickDateLastUsed.Checked = False
        Me.dtPickDateLastUsed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickDateLastUsed.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickDateLastUsed.Location = New System.Drawing.Point(152, 16)
        Me.dtPickDateLastUsed.Name = "dtPickDateLastUsed"
        Me.dtPickDateLastUsed.ShowCheckBox = True
        Me.dtPickDateLastUsed.Size = New System.Drawing.Size(104, 21)
        Me.dtPickDateLastUsed.TabIndex = 2
        Me.dtPickDateLastUsed.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'pnlClosureReport
        '
        Me.pnlClosureReport.Controls.Add(Me.lblClosureReportHead)
        Me.pnlClosureReport.Controls.Add(Me.lblClosureReportDisplay)
        Me.pnlClosureReport.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlClosureReport.Location = New System.Drawing.Point(3, 1080)
        Me.pnlClosureReport.Name = "pnlClosureReport"
        Me.pnlClosureReport.Size = New System.Drawing.Size(954, 24)
        Me.pnlClosureReport.TabIndex = 199
        '
        'lblClosureReportHead
        '
        Me.lblClosureReportHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblClosureReportHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblClosureReportHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblClosureReportHead.Location = New System.Drawing.Point(16, 0)
        Me.lblClosureReportHead.Name = "lblClosureReportHead"
        Me.lblClosureReportHead.Size = New System.Drawing.Size(938, 24)
        Me.lblClosureReportHead.TabIndex = 1
        Me.lblClosureReportHead.Text = "Closure Report"
        Me.lblClosureReportHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblClosureReportDisplay
        '
        Me.lblClosureReportDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblClosureReportDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblClosureReportDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblClosureReportDisplay.Name = "lblClosureReportDisplay"
        Me.lblClosureReportDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblClosureReportDisplay.TabIndex = 0
        Me.lblClosureReportDisplay.Text = "-"
        Me.lblClosureReportDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlChecklistDetails
        '
        Me.pnlChecklistDetails.Controls.Add(Me.lblDueBy)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickDueBy)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures1)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures1)
        Me.pnlChecklistDetails.Controls.Add(Me.lblDateClosed)
        Me.pnlChecklistDetails.Controls.Add(Me.lblOpen)
        Me.pnlChecklistDetails.Controls.Add(Me.lblDateClosed1)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures2)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures12)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures3)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures5)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures4)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures6)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures9)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures8)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures7)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures11)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures10)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures13)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures14)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures15)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures16)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures17)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures18)
        Me.pnlChecklistDetails.Controls.Add(Me.Label1)
        Me.pnlChecklistDetails.Controls.Add(Me.dtPickClosures19)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures2)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures3)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures4)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures5)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures6)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures7)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures8)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures9)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures10)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures11)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures12)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures13)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures14)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures15)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures16)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures17)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures18)
        Me.pnlChecklistDetails.Controls.Add(Me.chkClosures19)
        Me.pnlChecklistDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlChecklistDetails.Location = New System.Drawing.Point(3, 600)
        Me.pnlChecklistDetails.Name = "pnlChecklistDetails"
        Me.pnlChecklistDetails.Size = New System.Drawing.Size(954, 480)
        Me.pnlChecklistDetails.TabIndex = 20
        '
        'lblDueBy
        '
        Me.lblDueBy.Location = New System.Drawing.Point(48, 16)
        Me.lblDueBy.Name = "lblDueBy"
        Me.lblDueBy.Size = New System.Drawing.Size(48, 23)
        Me.lblDueBy.TabIndex = 3
        Me.lblDueBy.Text = "Due By:"
        Me.lblDueBy.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dtPickDueBy
        '
        Me.dtPickDueBy.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickDueBy.Checked = False
        Me.dtPickDueBy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickDueBy.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickDueBy.Location = New System.Drawing.Point(104, 16)
        Me.dtPickDueBy.Name = "dtPickDueBy"
        Me.dtPickDueBy.ShowCheckBox = True
        Me.dtPickDueBy.Size = New System.Drawing.Size(120, 21)
        Me.dtPickDueBy.TabIndex = 21
        Me.dtPickDueBy.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'chkClosures1
        '
        Me.chkClosures1.Location = New System.Drawing.Point(48, 72)
        Me.chkClosures1.Name = "chkClosures1"
        Me.chkClosures1.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures1.TabIndex = 22
        Me.chkClosures1.Tag = "1"
        '
        'dtPickClosures1
        '
        Me.dtPickClosures1.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures1.Checked = False
        Me.dtPickClosures1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures1.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures1.Location = New System.Drawing.Point(312, 80)
        Me.dtPickClosures1.Name = "dtPickClosures1"
        Me.dtPickClosures1.ShowCheckBox = True
        Me.dtPickClosures1.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures1.TabIndex = 23
        Me.dtPickClosures1.Tag = "1"
        Me.dtPickClosures1.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'lblDateClosed
        '
        Me.lblDateClosed.Location = New System.Drawing.Point(312, 48)
        Me.lblDateClosed.Name = "lblDateClosed"
        Me.lblDateClosed.Size = New System.Drawing.Size(72, 24)
        Me.lblDateClosed.TabIndex = 3
        Me.lblDateClosed.Text = "Date Closed"
        '
        'lblOpen
        '
        Me.lblOpen.Location = New System.Drawing.Point(32, 48)
        Me.lblOpen.Name = "lblOpen"
        Me.lblOpen.Size = New System.Drawing.Size(40, 24)
        Me.lblOpen.TabIndex = 3
        Me.lblOpen.Text = "Open"
        '
        'lblDateClosed1
        '
        Me.lblDateClosed1.Location = New System.Drawing.Point(728, 48)
        Me.lblDateClosed1.Name = "lblDateClosed1"
        Me.lblDateClosed1.Size = New System.Drawing.Size(72, 32)
        Me.lblDateClosed1.TabIndex = 3
        Me.lblDateClosed1.Text = "Date Closed"
        '
        'dtPickClosures2
        '
        Me.dtPickClosures2.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures2.Checked = False
        Me.dtPickClosures2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures2.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures2.Location = New System.Drawing.Point(312, 120)
        Me.dtPickClosures2.Name = "dtPickClosures2"
        Me.dtPickClosures2.ShowCheckBox = True
        Me.dtPickClosures2.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures2.TabIndex = 25
        Me.dtPickClosures2.Tag = "2"
        Me.dtPickClosures2.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures12
        '
        Me.dtPickClosures12.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures12.Checked = False
        Me.dtPickClosures12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures12.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures12.Location = New System.Drawing.Point(728, 120)
        Me.dtPickClosures12.Name = "dtPickClosures12"
        Me.dtPickClosures12.ShowCheckBox = True
        Me.dtPickClosures12.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures12.TabIndex = 45
        Me.dtPickClosures12.Tag = "12"
        Me.dtPickClosures12.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures3
        '
        Me.dtPickClosures3.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures3.Checked = False
        Me.dtPickClosures3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures3.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures3.Location = New System.Drawing.Point(312, 160)
        Me.dtPickClosures3.Name = "dtPickClosures3"
        Me.dtPickClosures3.ShowCheckBox = True
        Me.dtPickClosures3.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures3.TabIndex = 27
        Me.dtPickClosures3.Tag = "3"
        Me.dtPickClosures3.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures5
        '
        Me.dtPickClosures5.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures5.Checked = False
        Me.dtPickClosures5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures5.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures5.Location = New System.Drawing.Point(312, 240)
        Me.dtPickClosures5.Name = "dtPickClosures5"
        Me.dtPickClosures5.ShowCheckBox = True
        Me.dtPickClosures5.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures5.TabIndex = 31
        Me.dtPickClosures5.Tag = "5"
        Me.dtPickClosures5.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures4
        '
        Me.dtPickClosures4.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures4.Checked = False
        Me.dtPickClosures4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures4.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures4.Location = New System.Drawing.Point(312, 200)
        Me.dtPickClosures4.Name = "dtPickClosures4"
        Me.dtPickClosures4.ShowCheckBox = True
        Me.dtPickClosures4.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures4.TabIndex = 29
        Me.dtPickClosures4.Tag = "4"
        Me.dtPickClosures4.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures6
        '
        Me.dtPickClosures6.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures6.Checked = False
        Me.dtPickClosures6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures6.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures6.Location = New System.Drawing.Point(312, 280)
        Me.dtPickClosures6.Name = "dtPickClosures6"
        Me.dtPickClosures6.ShowCheckBox = True
        Me.dtPickClosures6.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures6.TabIndex = 33
        Me.dtPickClosures6.Tag = "6"
        Me.dtPickClosures6.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures9
        '
        Me.dtPickClosures9.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures9.Checked = False
        Me.dtPickClosures9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures9.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures9.Location = New System.Drawing.Point(312, 400)
        Me.dtPickClosures9.Name = "dtPickClosures9"
        Me.dtPickClosures9.ShowCheckBox = True
        Me.dtPickClosures9.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures9.TabIndex = 39
        Me.dtPickClosures9.Tag = "9"
        Me.dtPickClosures9.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures8
        '
        Me.dtPickClosures8.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures8.Checked = False
        Me.dtPickClosures8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures8.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures8.Location = New System.Drawing.Point(312, 360)
        Me.dtPickClosures8.Name = "dtPickClosures8"
        Me.dtPickClosures8.ShowCheckBox = True
        Me.dtPickClosures8.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures8.TabIndex = 37
        Me.dtPickClosures8.Tag = "8"
        Me.dtPickClosures8.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures7
        '
        Me.dtPickClosures7.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures7.Checked = False
        Me.dtPickClosures7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures7.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures7.Location = New System.Drawing.Point(312, 320)
        Me.dtPickClosures7.Name = "dtPickClosures7"
        Me.dtPickClosures7.ShowCheckBox = True
        Me.dtPickClosures7.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures7.TabIndex = 35
        Me.dtPickClosures7.Tag = "7"
        Me.dtPickClosures7.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures11
        '
        Me.dtPickClosures11.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures11.Checked = False
        Me.dtPickClosures11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures11.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures11.Location = New System.Drawing.Point(728, 80)
        Me.dtPickClosures11.Name = "dtPickClosures11"
        Me.dtPickClosures11.ShowCheckBox = True
        Me.dtPickClosures11.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures11.TabIndex = 43
        Me.dtPickClosures11.Tag = "11"
        Me.dtPickClosures11.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures10
        '
        Me.dtPickClosures10.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures10.Checked = False
        Me.dtPickClosures10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures10.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures10.Location = New System.Drawing.Point(312, 440)
        Me.dtPickClosures10.Name = "dtPickClosures10"
        Me.dtPickClosures10.ShowCheckBox = True
        Me.dtPickClosures10.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures10.TabIndex = 41
        Me.dtPickClosures10.Tag = "10"
        Me.dtPickClosures10.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures13
        '
        Me.dtPickClosures13.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures13.Checked = False
        Me.dtPickClosures13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures13.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures13.Location = New System.Drawing.Point(728, 160)
        Me.dtPickClosures13.Name = "dtPickClosures13"
        Me.dtPickClosures13.ShowCheckBox = True
        Me.dtPickClosures13.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures13.TabIndex = 47
        Me.dtPickClosures13.Tag = "13"
        Me.dtPickClosures13.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures14
        '
        Me.dtPickClosures14.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures14.Checked = False
        Me.dtPickClosures14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures14.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures14.Location = New System.Drawing.Point(728, 200)
        Me.dtPickClosures14.Name = "dtPickClosures14"
        Me.dtPickClosures14.ShowCheckBox = True
        Me.dtPickClosures14.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures14.TabIndex = 49
        Me.dtPickClosures14.Tag = "14"
        Me.dtPickClosures14.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures15
        '
        Me.dtPickClosures15.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures15.Checked = False
        Me.dtPickClosures15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures15.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures15.Location = New System.Drawing.Point(728, 240)
        Me.dtPickClosures15.Name = "dtPickClosures15"
        Me.dtPickClosures15.ShowCheckBox = True
        Me.dtPickClosures15.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures15.TabIndex = 51
        Me.dtPickClosures15.Tag = "15"
        Me.dtPickClosures15.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures16
        '
        Me.dtPickClosures16.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures16.Checked = False
        Me.dtPickClosures16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures16.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures16.Location = New System.Drawing.Point(728, 280)
        Me.dtPickClosures16.Name = "dtPickClosures16"
        Me.dtPickClosures16.ShowCheckBox = True
        Me.dtPickClosures16.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures16.TabIndex = 53
        Me.dtPickClosures16.Tag = "16"
        Me.dtPickClosures16.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures17
        '
        Me.dtPickClosures17.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures17.Checked = False
        Me.dtPickClosures17.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures17.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures17.Location = New System.Drawing.Point(728, 320)
        Me.dtPickClosures17.Name = "dtPickClosures17"
        Me.dtPickClosures17.ShowCheckBox = True
        Me.dtPickClosures17.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures17.TabIndex = 55
        Me.dtPickClosures17.Tag = "17"
        Me.dtPickClosures17.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'dtPickClosures18
        '
        Me.dtPickClosures18.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures18.Checked = False
        Me.dtPickClosures18.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures18.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures18.Location = New System.Drawing.Point(728, 360)
        Me.dtPickClosures18.Name = "dtPickClosures18"
        Me.dtPickClosures18.ShowCheckBox = True
        Me.dtPickClosures18.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures18.TabIndex = 56
        Me.dtPickClosures18.Tag = "18"
        Me.dtPickClosures18.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(432, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 24)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Open"
        '
        'dtPickClosures19
        '
        Me.dtPickClosures19.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures19.Checked = False
        Me.dtPickClosures19.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickClosures19.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickClosures19.Location = New System.Drawing.Point(728, 400)
        Me.dtPickClosures19.Name = "dtPickClosures19"
        Me.dtPickClosures19.ShowCheckBox = True
        Me.dtPickClosures19.Size = New System.Drawing.Size(104, 21)
        Me.dtPickClosures19.TabIndex = 56
        Me.dtPickClosures19.Tag = "19"
        Me.dtPickClosures19.Value = New Date(2004, 11, 15, 0, 0, 0, 0)
        '
        'chkClosures2
        '
        Me.chkClosures2.Location = New System.Drawing.Point(48, 112)
        Me.chkClosures2.Name = "chkClosures2"
        Me.chkClosures2.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures2.TabIndex = 22
        Me.chkClosures2.Tag = "2"
        '
        'chkClosures3
        '
        Me.chkClosures3.Location = New System.Drawing.Point(48, 152)
        Me.chkClosures3.Name = "chkClosures3"
        Me.chkClosures3.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures3.TabIndex = 22
        Me.chkClosures3.Tag = "3"
        '
        'chkClosures4
        '
        Me.chkClosures4.Location = New System.Drawing.Point(48, 192)
        Me.chkClosures4.Name = "chkClosures4"
        Me.chkClosures4.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures4.TabIndex = 22
        Me.chkClosures4.Tag = "4"
        '
        'chkClosures5
        '
        Me.chkClosures5.Location = New System.Drawing.Point(48, 232)
        Me.chkClosures5.Name = "chkClosures5"
        Me.chkClosures5.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures5.TabIndex = 22
        Me.chkClosures5.Tag = "5"
        '
        'chkClosures6
        '
        Me.chkClosures6.Location = New System.Drawing.Point(48, 272)
        Me.chkClosures6.Name = "chkClosures6"
        Me.chkClosures6.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures6.TabIndex = 22
        Me.chkClosures6.Tag = "6"
        '
        'chkClosures7
        '
        Me.chkClosures7.Location = New System.Drawing.Point(48, 312)
        Me.chkClosures7.Name = "chkClosures7"
        Me.chkClosures7.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures7.TabIndex = 22
        Me.chkClosures7.Tag = "7"
        '
        'chkClosures8
        '
        Me.chkClosures8.Location = New System.Drawing.Point(48, 352)
        Me.chkClosures8.Name = "chkClosures8"
        Me.chkClosures8.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures8.TabIndex = 22
        Me.chkClosures8.Tag = "8"
        '
        'chkClosures9
        '
        Me.chkClosures9.Location = New System.Drawing.Point(48, 392)
        Me.chkClosures9.Name = "chkClosures9"
        Me.chkClosures9.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures9.TabIndex = 22
        Me.chkClosures9.Tag = "9"
        '
        'chkClosures10
        '
        Me.chkClosures10.Location = New System.Drawing.Point(48, 432)
        Me.chkClosures10.Name = "chkClosures10"
        Me.chkClosures10.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures10.TabIndex = 22
        Me.chkClosures10.Tag = "10"
        '
        'chkClosures11
        '
        Me.chkClosures11.Location = New System.Drawing.Point(448, 72)
        Me.chkClosures11.Name = "chkClosures11"
        Me.chkClosures11.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures11.TabIndex = 22
        Me.chkClosures11.Tag = "11"
        '
        'chkClosures12
        '
        Me.chkClosures12.Location = New System.Drawing.Point(448, 112)
        Me.chkClosures12.Name = "chkClosures12"
        Me.chkClosures12.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures12.TabIndex = 22
        Me.chkClosures12.Tag = "12"
        '
        'chkClosures13
        '
        Me.chkClosures13.Location = New System.Drawing.Point(448, 152)
        Me.chkClosures13.Name = "chkClosures13"
        Me.chkClosures13.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures13.TabIndex = 22
        Me.chkClosures13.Tag = "13"
        '
        'chkClosures14
        '
        Me.chkClosures14.Location = New System.Drawing.Point(448, 192)
        Me.chkClosures14.Name = "chkClosures14"
        Me.chkClosures14.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures14.TabIndex = 22
        Me.chkClosures14.Tag = "14"
        '
        'chkClosures15
        '
        Me.chkClosures15.Location = New System.Drawing.Point(448, 232)
        Me.chkClosures15.Name = "chkClosures15"
        Me.chkClosures15.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures15.TabIndex = 22
        Me.chkClosures15.Tag = "15"
        '
        'chkClosures16
        '
        Me.chkClosures16.Location = New System.Drawing.Point(448, 272)
        Me.chkClosures16.Name = "chkClosures16"
        Me.chkClosures16.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures16.TabIndex = 22
        Me.chkClosures16.Tag = "16"
        '
        'chkClosures17
        '
        Me.chkClosures17.Location = New System.Drawing.Point(448, 312)
        Me.chkClosures17.Name = "chkClosures17"
        Me.chkClosures17.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures17.TabIndex = 22
        Me.chkClosures17.Tag = "17"
        '
        'chkClosures18
        '
        Me.chkClosures18.Location = New System.Drawing.Point(448, 352)
        Me.chkClosures18.Name = "chkClosures18"
        Me.chkClosures18.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures18.TabIndex = 22
        Me.chkClosures18.Tag = "18"
        '
        'chkClosures19
        '
        Me.chkClosures19.Location = New System.Drawing.Point(448, 392)
        Me.chkClosures19.Name = "chkClosures19"
        Me.chkClosures19.Size = New System.Drawing.Size(240, 32)
        Me.chkClosures19.TabIndex = 22
        Me.chkClosures19.Tag = "19"
        '
        'pnlChecklist
        '
        Me.pnlChecklist.Controls.Add(Me.lblChecklistHead)
        Me.pnlChecklist.Controls.Add(Me.lblChecklistDisplay)
        Me.pnlChecklist.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlChecklist.Location = New System.Drawing.Point(3, 576)
        Me.pnlChecklist.Name = "pnlChecklist"
        Me.pnlChecklist.Size = New System.Drawing.Size(954, 24)
        Me.pnlChecklist.TabIndex = 192
        '
        'lblChecklistHead
        '
        Me.lblChecklistHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblChecklistHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblChecklistHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblChecklistHead.Location = New System.Drawing.Point(16, 0)
        Me.lblChecklistHead.Name = "lblChecklistHead"
        Me.lblChecklistHead.Size = New System.Drawing.Size(938, 24)
        Me.lblChecklistHead.TabIndex = 1
        Me.lblChecklistHead.Text = "Checklist"
        Me.lblChecklistHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblChecklistDisplay
        '
        Me.lblChecklistDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblChecklistDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblChecklistDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblChecklistDisplay.Name = "lblChecklistDisplay"
        Me.lblChecklistDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblChecklistDisplay.TabIndex = 0
        Me.lblChecklistDisplay.Text = "-"
        Me.lblChecklistDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlAnalysisDetails
        '
        Me.pnlAnalysisDetails.Controls.Add(Me.dGridAnalysis)
        Me.pnlAnalysisDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAnalysisDetails.Location = New System.Drawing.Point(3, 440)
        Me.pnlAnalysisDetails.Name = "pnlAnalysisDetails"
        Me.pnlAnalysisDetails.Size = New System.Drawing.Size(954, 136)
        Me.pnlAnalysisDetails.TabIndex = 18
        '
        'dGridAnalysis
        '
        Me.dGridAnalysis.Cursor = System.Windows.Forms.Cursors.Default
        Me.dGridAnalysis.Location = New System.Drawing.Point(40, 8)
        Me.dGridAnalysis.Name = "dGridAnalysis"
        Me.dGridAnalysis.Size = New System.Drawing.Size(880, 112)
        Me.dGridAnalysis.TabIndex = 19
        '
        'pnlAnalysis
        '
        Me.pnlAnalysis.Controls.Add(Me.lblAnalysisHead)
        Me.pnlAnalysis.Controls.Add(Me.lblAnalysisDisplay)
        Me.pnlAnalysis.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAnalysis.Location = New System.Drawing.Point(3, 416)
        Me.pnlAnalysis.Name = "pnlAnalysis"
        Me.pnlAnalysis.Size = New System.Drawing.Size(954, 24)
        Me.pnlAnalysis.TabIndex = 190
        '
        'lblAnalysisHead
        '
        Me.lblAnalysisHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblAnalysisHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblAnalysisHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblAnalysisHead.Location = New System.Drawing.Point(16, 0)
        Me.lblAnalysisHead.Name = "lblAnalysisHead"
        Me.lblAnalysisHead.Size = New System.Drawing.Size(938, 24)
        Me.lblAnalysisHead.TabIndex = 1
        Me.lblAnalysisHead.Text = "Analysis"
        Me.lblAnalysisHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAnalysisDisplay
        '
        Me.lblAnalysisDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAnalysisDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblAnalysisDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblAnalysisDisplay.Name = "lblAnalysisDisplay"
        Me.lblAnalysisDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblAnalysisDisplay.TabIndex = 0
        Me.lblAnalysisDisplay.Text = "-"
        Me.lblAnalysisDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PnlTanksPipesDetails
        '
        Me.PnlTanksPipesDetails.Controls.Add(Me.udPreviousSubstance)
        Me.PnlTanksPipesDetails.Controls.Add(Me.btnEvtTankCollapse)
        Me.PnlTanksPipesDetails.Controls.Add(Me.chkShowAllTanksPipes)
        Me.PnlTanksPipesDetails.Controls.Add(Me.ugTankandPipes)
        Me.PnlTanksPipesDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.PnlTanksPipesDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlTanksPipesDetails.Location = New System.Drawing.Point(3, 168)
        Me.PnlTanksPipesDetails.Name = "PnlTanksPipesDetails"
        Me.PnlTanksPipesDetails.Size = New System.Drawing.Size(954, 248)
        Me.PnlTanksPipesDetails.TabIndex = 14
        '
        'udPreviousSubstance
        '
        Me.udPreviousSubstance.Cursor = System.Windows.Forms.Cursors.Default
        Me.udPreviousSubstance.DisplayMember = ""
        Me.udPreviousSubstance.Location = New System.Drawing.Point(624, 128)
        Me.udPreviousSubstance.Name = "udPreviousSubstance"
        Me.udPreviousSubstance.Size = New System.Drawing.Size(160, 80)
        Me.udPreviousSubstance.TabIndex = 219
        Me.udPreviousSubstance.Text = "Previous Substance"
        Me.udPreviousSubstance.ValueMember = ""
        Me.udPreviousSubstance.Visible = False
        '
        'btnEvtTankCollapse
        '
        Me.btnEvtTankCollapse.Location = New System.Drawing.Point(224, 0)
        Me.btnEvtTankCollapse.Name = "btnEvtTankCollapse"
        Me.btnEvtTankCollapse.Size = New System.Drawing.Size(80, 23)
        Me.btnEvtTankCollapse.TabIndex = 16
        Me.btnEvtTankCollapse.Text = "Collapse All"
        '
        'chkShowAllTanksPipes
        '
        Me.chkShowAllTanksPipes.Checked = True
        Me.chkShowAllTanksPipes.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowAllTanksPipes.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.chkShowAllTanksPipes.Location = New System.Drawing.Point(80, 0)
        Me.chkShowAllTanksPipes.Name = "chkShowAllTanksPipes"
        Me.chkShowAllTanksPipes.Size = New System.Drawing.Size(144, 21)
        Me.chkShowAllTanksPipes.TabIndex = 15
        Me.chkShowAllTanksPipes.Text = "Show All Tanks/Pipes"
        '
        'ugTankandPipes
        '
        Me.ugTankandPipes.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTankandPipes.Location = New System.Drawing.Point(42, 24)
        Me.ugTankandPipes.Name = "ugTankandPipes"
        Me.ugTankandPipes.Size = New System.Drawing.Size(880, 216)
        Me.ugTankandPipes.TabIndex = 17
        Me.ugTankandPipes.Text = "Tank(s) / Pipe(s)"
        '
        'pnlTanksPipes
        '
        Me.pnlTanksPipes.Controls.Add(Me.lblTanksPipesHead)
        Me.pnlTanksPipes.Controls.Add(Me.lblTanksPipesDisplay)
        Me.pnlTanksPipes.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTanksPipes.Location = New System.Drawing.Point(3, 144)
        Me.pnlTanksPipes.Name = "pnlTanksPipes"
        Me.pnlTanksPipes.Size = New System.Drawing.Size(954, 24)
        Me.pnlTanksPipes.TabIndex = 9
        '
        'lblTanksPipesHead
        '
        Me.lblTanksPipesHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTanksPipesHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTanksPipesHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTanksPipesHead.Location = New System.Drawing.Point(16, 0)
        Me.lblTanksPipesHead.Name = "lblTanksPipesHead"
        Me.lblTanksPipesHead.Size = New System.Drawing.Size(938, 24)
        Me.lblTanksPipesHead.TabIndex = 1
        Me.lblTanksPipesHead.Text = "Tanks / Pipes"
        Me.lblTanksPipesHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTanksPipesDisplay
        '
        Me.lblTanksPipesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTanksPipesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTanksPipesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblTanksPipesDisplay.Name = "lblTanksPipesDisplay"
        Me.lblTanksPipesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblTanksPipesDisplay.TabIndex = 0
        Me.lblTanksPipesDisplay.Text = "-"
        Me.lblTanksPipesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PnlNoticeOfInterestDetails
        '
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.lblNOISearchCompany)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.txtNOILicensee)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.txtNOICompany)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.cmbVerbalWaiver)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.lblCertifiedContractor)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.cmbFillMaterial)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.lblFillMaterial)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.lblCompany)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.lblScheduledDate)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.dtPickScheduledDate)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.dtPickReceived)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.lblReceived)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.btnProcessNOI)
        Me.PnlNoticeOfInterestDetails.Controls.Add(Me.btnProcessNOIEnvelopes)
        Me.PnlNoticeOfInterestDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.PnlNoticeOfInterestDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PnlNoticeOfInterestDetails.Location = New System.Drawing.Point(3, 27)
        Me.PnlNoticeOfInterestDetails.Name = "PnlNoticeOfInterestDetails"
        Me.PnlNoticeOfInterestDetails.Size = New System.Drawing.Size(954, 117)
        Me.PnlNoticeOfInterestDetails.TabIndex = 4
        '
        'lblNOISearchCompany
        '
        Me.lblNOISearchCompany.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNOISearchCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNOISearchCompany.ForeColor = System.Drawing.Color.MediumTurquoise
        Me.lblNOISearchCompany.Location = New System.Drawing.Point(312, 48)
        Me.lblNOISearchCompany.Name = "lblNOISearchCompany"
        Me.lblNOISearchCompany.Size = New System.Drawing.Size(16, 22)
        Me.lblNOISearchCompany.TabIndex = 225
        Me.lblNOISearchCompany.Text = "?"
        Me.lblNOISearchCompany.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtNOILicensee
        '
        Me.txtNOILicensee.Location = New System.Drawing.Point(136, 48)
        Me.txtNOILicensee.Name = "txtNOILicensee"
        Me.txtNOILicensee.Size = New System.Drawing.Size(168, 21)
        Me.txtNOILicensee.TabIndex = 224
        Me.txtNOILicensee.Text = ""
        '
        'txtNOICompany
        '
        Me.txtNOICompany.Location = New System.Drawing.Point(136, 80)
        Me.txtNOICompany.Name = "txtNOICompany"
        Me.txtNOICompany.Size = New System.Drawing.Size(168, 21)
        Me.txtNOICompany.TabIndex = 223
        Me.txtNOICompany.Text = ""
        '
        'cmbVerbalWaiver
        '
        Me.cmbVerbalWaiver.Location = New System.Drawing.Point(376, 80)
        Me.cmbVerbalWaiver.Name = "cmbVerbalWaiver"
        Me.cmbVerbalWaiver.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.cmbVerbalWaiver.Size = New System.Drawing.Size(104, 20)
        Me.cmbVerbalWaiver.TabIndex = 5
        Me.cmbVerbalWaiver.Text = "Verbal Waiver"
        Me.cmbVerbalWaiver.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCertifiedContractor
        '
        Me.lblCertifiedContractor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCertifiedContractor.Location = New System.Drawing.Point(8, 48)
        Me.lblCertifiedContractor.Name = "lblCertifiedContractor"
        Me.lblCertifiedContractor.Size = New System.Drawing.Size(120, 23)
        Me.lblCertifiedContractor.TabIndex = 185
        Me.lblCertifiedContractor.Text = "Certified Contractor:"
        Me.lblCertifiedContractor.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbFillMaterial
        '
        Me.cmbFillMaterial.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFillMaterial.DropDownWidth = 144
        Me.cmbFillMaterial.Enabled = False
        Me.cmbFillMaterial.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbFillMaterial.ItemHeight = 15
        Me.cmbFillMaterial.Location = New System.Drawing.Point(465, 48)
        Me.cmbFillMaterial.Name = "cmbFillMaterial"
        Me.cmbFillMaterial.Size = New System.Drawing.Size(144, 23)
        Me.cmbFillMaterial.TabIndex = 4
        '
        'lblFillMaterial
        '
        Me.lblFillMaterial.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFillMaterial.Location = New System.Drawing.Point(384, 48)
        Me.lblFillMaterial.Name = "lblFillMaterial"
        Me.lblFillMaterial.Size = New System.Drawing.Size(72, 17)
        Me.lblFillMaterial.TabIndex = 183
        Me.lblFillMaterial.Text = "Fill Material:"
        Me.lblFillMaterial.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCompany
        '
        Me.lblCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompany.Location = New System.Drawing.Point(56, 80)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(72, 23)
        Me.lblCompany.TabIndex = 180
        Me.lblCompany.Text = "Company:"
        Me.lblCompany.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblScheduledDate
        '
        Me.lblScheduledDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblScheduledDate.Location = New System.Drawing.Point(368, 16)
        Me.lblScheduledDate.Name = "lblScheduledDate"
        Me.lblScheduledDate.Size = New System.Drawing.Size(96, 17)
        Me.lblScheduledDate.TabIndex = 174
        Me.lblScheduledDate.Text = "Scheduled Date:"
        '
        'dtPickScheduledDate
        '
        Me.dtPickScheduledDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickScheduledDate.Checked = False
        Me.dtPickScheduledDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickScheduledDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickScheduledDate.Location = New System.Drawing.Point(465, 16)
        Me.dtPickScheduledDate.Name = "dtPickScheduledDate"
        Me.dtPickScheduledDate.ShowCheckBox = True
        Me.dtPickScheduledDate.Size = New System.Drawing.Size(104, 21)
        Me.dtPickScheduledDate.TabIndex = 3
        Me.dtPickScheduledDate.Value = New Date(2005, 3, 29, 0, 0, 0, 0)
        '
        'dtPickReceived
        '
        Me.dtPickReceived.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickReceived.Checked = False
        Me.dtPickReceived.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtPickReceived.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickReceived.Location = New System.Drawing.Point(128, 16)
        Me.dtPickReceived.Name = "dtPickReceived"
        Me.dtPickReceived.ShowCheckBox = True
        Me.dtPickReceived.Size = New System.Drawing.Size(104, 21)
        Me.dtPickReceived.TabIndex = 0
        Me.dtPickReceived.Value = New Date(2005, 3, 29, 0, 0, 0, 0)
        '
        'lblReceived
        '
        Me.lblReceived.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReceived.Location = New System.Drawing.Point(64, 16)
        Me.lblReceived.Name = "lblReceived"
        Me.lblReceived.Size = New System.Drawing.Size(64, 17)
        Me.lblReceived.TabIndex = 174
        Me.lblReceived.Text = "Received:"
        '
        'btnProcessNOI
        '
        Me.btnProcessNOI.Enabled = False
        Me.btnProcessNOI.Location = New System.Drawing.Point(744, 8)
        Me.btnProcessNOI.Name = "btnProcessNOI"
        Me.btnProcessNOI.Size = New System.Drawing.Size(184, 24)
        Me.btnProcessNOI.TabIndex = 6
        Me.btnProcessNOI.Text = "Process NOI"
        '
        'btnProcessNOIEnvelopes
        '
        Me.btnProcessNOIEnvelopes.Location = New System.Drawing.Point(744, 40)
        Me.btnProcessNOIEnvelopes.Name = "btnProcessNOIEnvelopes"
        Me.btnProcessNOIEnvelopes.Size = New System.Drawing.Size(184, 24)
        Me.btnProcessNOIEnvelopes.TabIndex = 6
        Me.btnProcessNOIEnvelopes.Text = "Envelopes"
        '
        'pnlNoticeOfInterest
        '
        Me.pnlNoticeOfInterest.Controls.Add(Me.lblNoticeOfInterestHead)
        Me.pnlNoticeOfInterest.Controls.Add(Me.lblNoticeOfInterestDisplay)
        Me.pnlNoticeOfInterest.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlNoticeOfInterest.Location = New System.Drawing.Point(3, 3)
        Me.pnlNoticeOfInterest.Name = "pnlNoticeOfInterest"
        Me.pnlNoticeOfInterest.Size = New System.Drawing.Size(954, 24)
        Me.pnlNoticeOfInterest.TabIndex = 7
        '
        'lblNoticeOfInterestHead
        '
        Me.lblNoticeOfInterestHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblNoticeOfInterestHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblNoticeOfInterestHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblNoticeOfInterestHead.Location = New System.Drawing.Point(16, 0)
        Me.lblNoticeOfInterestHead.Name = "lblNoticeOfInterestHead"
        Me.lblNoticeOfInterestHead.Size = New System.Drawing.Size(938, 24)
        Me.lblNoticeOfInterestHead.TabIndex = 1
        Me.lblNoticeOfInterestHead.Text = "Notice of Intent"
        Me.lblNoticeOfInterestHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblNoticeOfInterestDisplay
        '
        Me.lblNoticeOfInterestDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNoticeOfInterestDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoticeOfInterestDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblNoticeOfInterestDisplay.Name = "lblNoticeOfInterestDisplay"
        Me.lblNoticeOfInterestDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblNoticeOfInterestDisplay.TabIndex = 0
        Me.lblNoticeOfInterestDisplay.Text = "-"
        Me.lblNoticeOfInterestDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlLUSTEventBottom
        '
        Me.pnlLUSTEventBottom.Controls.Add(Me.btnDeleteClosure)
        Me.pnlLUSTEventBottom.Controls.Add(Me.btnClosureComments)
        Me.pnlLUSTEventBottom.Controls.Add(Me.btnClosureCancel)
        Me.pnlLUSTEventBottom.Controls.Add(Me.btnSaveClosure)
        Me.pnlLUSTEventBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLUSTEventBottom.Location = New System.Drawing.Point(0, 628)
        Me.pnlLUSTEventBottom.Name = "pnlLUSTEventBottom"
        Me.pnlLUSTEventBottom.Size = New System.Drawing.Size(964, 40)
        Me.pnlLUSTEventBottom.TabIndex = 66
        '
        'btnDeleteClosure
        '
        Me.btnDeleteClosure.Location = New System.Drawing.Point(440, 8)
        Me.btnDeleteClosure.Name = "btnDeleteClosure"
        Me.btnDeleteClosure.Size = New System.Drawing.Size(104, 26)
        Me.btnDeleteClosure.TabIndex = 69
        Me.btnDeleteClosure.Text = "Delete Closure"
        '
        'btnClosureComments
        '
        Me.btnClosureComments.Location = New System.Drawing.Point(552, 8)
        Me.btnClosureComments.Name = "btnClosureComments"
        Me.btnClosureComments.Size = New System.Drawing.Size(104, 26)
        Me.btnClosureComments.TabIndex = 70
        Me.btnClosureComments.Text = "Comments"
        '
        'btnClosureCancel
        '
        Me.btnClosureCancel.Enabled = False
        Me.btnClosureCancel.Location = New System.Drawing.Point(328, 8)
        Me.btnClosureCancel.Name = "btnClosureCancel"
        Me.btnClosureCancel.Size = New System.Drawing.Size(104, 26)
        Me.btnClosureCancel.TabIndex = 68
        Me.btnClosureCancel.Text = "Cancel"
        '
        'btnSaveClosure
        '
        Me.btnSaveClosure.Enabled = False
        Me.btnSaveClosure.Location = New System.Drawing.Point(216, 8)
        Me.btnSaveClosure.Name = "btnSaveClosure"
        Me.btnSaveClosure.Size = New System.Drawing.Size(104, 26)
        Me.btnSaveClosure.TabIndex = 67
        Me.btnSaveClosure.Text = "Save Closure"
        '
        'pnlNoticeOfInterestHeader
        '
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lnkLblNextClosure)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lnkLblPrevClosure)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.btnReopen)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lblNOIReceived)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lblClosureCountValue)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lblClosureIDValue)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lblClosureID)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.cmbClosureType)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lblProjectManager)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lblStatus)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.lblStatusValue)
        Me.pnlNoticeOfInterestHeader.Controls.Add(Me.GroupBox1)
        Me.pnlNoticeOfInterestHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlNoticeOfInterestHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlNoticeOfInterestHeader.Name = "pnlNoticeOfInterestHeader"
        Me.pnlNoticeOfInterestHeader.Size = New System.Drawing.Size(964, 64)
        Me.pnlNoticeOfInterestHeader.TabIndex = 0
        '
        'lnkLblNextClosure
        '
        Me.lnkLblNextClosure.Location = New System.Drawing.Point(88, 8)
        Me.lnkLblNextClosure.Name = "lnkLblNextClosure"
        Me.lnkLblNextClosure.Size = New System.Drawing.Size(56, 16)
        Me.lnkLblNextClosure.TabIndex = 1044
        Me.lnkLblNextClosure.TabStop = True
        Me.lnkLblNextClosure.Text = "Next>>"
        '
        'lnkLblPrevClosure
        '
        Me.lnkLblPrevClosure.AutoSize = True
        Me.lnkLblPrevClosure.Location = New System.Drawing.Point(8, 8)
        Me.lnkLblPrevClosure.Name = "lnkLblPrevClosure"
        Me.lnkLblPrevClosure.Size = New System.Drawing.Size(70, 17)
        Me.lnkLblPrevClosure.TabIndex = 1043
        Me.lnkLblPrevClosure.TabStop = True
        Me.lnkLblPrevClosure.Text = "<< Previous"
        '
        'btnReopen
        '
        Me.btnReopen.Enabled = False
        Me.btnReopen.Location = New System.Drawing.Point(761, 32)
        Me.btnReopen.Name = "btnReopen"
        Me.btnReopen.Size = New System.Drawing.Size(75, 24)
        Me.btnReopen.TabIndex = 214
        Me.btnReopen.Text = "Reopen"
        '
        'lblNOIReceived
        '
        Me.lblNOIReceived.Location = New System.Drawing.Point(359, 32)
        Me.lblNOIReceived.Name = "lblNOIReceived"
        Me.lblNOIReceived.Size = New System.Drawing.Size(88, 23)
        Me.lblNOIReceived.TabIndex = 213
        Me.lblNOIReceived.Text = "NOI Received:"
        '
        'lblClosureCountValue
        '
        Me.lblClosureCountValue.AutoSize = True
        Me.lblClosureCountValue.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblClosureCountValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosureCountValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblClosureCountValue.Location = New System.Drawing.Point(88, 32)
        Me.lblClosureCountValue.Name = "lblClosureCountValue"
        Me.lblClosureCountValue.Size = New System.Drawing.Size(39, 17)
        Me.lblClosureCountValue.TabIndex = 210
        Me.lblClosureCountValue.Text = "of ???"
        '
        'lblClosureIDValue
        '
        Me.lblClosureIDValue.AutoSize = True
        Me.lblClosureIDValue.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblClosureIDValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblClosureIDValue.Location = New System.Drawing.Point(73, 32)
        Me.lblClosureIDValue.Name = "lblClosureIDValue"
        Me.lblClosureIDValue.Size = New System.Drawing.Size(18, 17)
        Me.lblClosureIDValue.TabIndex = 209
        Me.lblClosureIDValue.Text = "00"
        Me.lblClosureIDValue.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblClosureID
        '
        Me.lblClosureID.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblClosureID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClosureID.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblClosureID.Location = New System.Drawing.Point(8, 32)
        Me.lblClosureID.Name = "lblClosureID"
        Me.lblClosureID.Size = New System.Drawing.Size(64, 17)
        Me.lblClosureID.TabIndex = 208
        Me.lblClosureID.Text = "Closure #:"
        '
        'cmbClosureType
        '
        Me.cmbClosureType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbClosureType.DropDownWidth = 128
        Me.cmbClosureType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbClosureType.ItemHeight = 15
        Me.cmbClosureType.Location = New System.Drawing.Point(202, 32)
        Me.cmbClosureType.Name = "cmbClosureType"
        Me.cmbClosureType.Size = New System.Drawing.Size(128, 23)
        Me.cmbClosureType.TabIndex = 0
        '
        'lblProjectManager
        '
        Me.lblProjectManager.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectManager.Location = New System.Drawing.Point(162, 32)
        Me.lblProjectManager.Name = "lblProjectManager"
        Me.lblProjectManager.Size = New System.Drawing.Size(40, 17)
        Me.lblProjectManager.TabIndex = 139
        Me.lblProjectManager.Text = "Type:"
        Me.lblProjectManager.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(589, 32)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(43, 23)
        Me.lblStatus.TabIndex = 213
        Me.lblStatus.Text = "Status:"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblStatusValue
        '
        Me.lblStatusValue.BackColor = System.Drawing.Color.LightGray
        Me.lblStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblStatusValue.Location = New System.Drawing.Point(632, 32)
        Me.lblStatusValue.Name = "lblStatusValue"
        Me.lblStatusValue.Size = New System.Drawing.Size(124, 23)
        Me.lblStatusValue.TabIndex = 213
        Me.lblStatusValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbtnNOIReceivedNo)
        Me.GroupBox1.Controls.Add(Me.rbtnNOIReceivedYes)
        Me.GroupBox1.Location = New System.Drawing.Point(447, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(112, 40)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'rbtnNOIReceivedNo
        '
        Me.rbtnNOIReceivedNo.Location = New System.Drawing.Point(64, 12)
        Me.rbtnNOIReceivedNo.Name = "rbtnNOIReceivedNo"
        Me.rbtnNOIReceivedNo.Size = New System.Drawing.Size(40, 24)
        Me.rbtnNOIReceivedNo.TabIndex = 3
        Me.rbtnNOIReceivedNo.Text = "No"
        '
        'rbtnNOIReceivedYes
        '
        Me.rbtnNOIReceivedYes.Location = New System.Drawing.Point(8, 12)
        Me.rbtnNOIReceivedYes.Name = "rbtnNOIReceivedYes"
        Me.rbtnNOIReceivedYes.Size = New System.Drawing.Size(48, 24)
        Me.rbtnNOIReceivedYes.TabIndex = 2
        Me.rbtnNOIReceivedYes.Text = "Yes"
        '
        'tbPageSummary
        '
        Me.tbPageSummary.AutoScroll = True
        Me.tbPageSummary.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageSummary.Controls.Add(Me.UCOwnerSummary)
        Me.tbPageSummary.Controls.Add(Me.pnlOwnerSummaryHeader)
        Me.tbPageSummary.Controls.Add(Me.Panel12)
        Me.tbPageSummary.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPageSummary.Location = New System.Drawing.Point(4, 22)
        Me.tbPageSummary.Name = "tbPageSummary"
        Me.tbPageSummary.Size = New System.Drawing.Size(964, 668)
        Me.tbPageSummary.TabIndex = 0
        Me.tbPageSummary.Text = "Owner Summary"
        Me.tbPageSummary.Visible = False
        '
        'UCOwnerSummary
        '
        Me.UCOwnerSummary.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerSummary.Location = New System.Drawing.Point(0, 16)
        Me.UCOwnerSummary.Name = "UCOwnerSummary"
        Me.UCOwnerSummary.Size = New System.Drawing.Size(952, 648)
        Me.UCOwnerSummary.TabIndex = 9
        '
        'pnlOwnerSummaryHeader
        '
        Me.pnlOwnerSummaryHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerSummaryHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerSummaryHeader.Name = "pnlOwnerSummaryHeader"
        Me.pnlOwnerSummaryHeader.Size = New System.Drawing.Size(952, 16)
        Me.pnlOwnerSummaryHeader.TabIndex = 8
        '
        'Panel12
        '
        Me.Panel12.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel12.DockPadding.Left = 10
        Me.Panel12.Location = New System.Drawing.Point(952, 0)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(8, 664)
        Me.Panel12.TabIndex = 6
        '
        'lblInViolationValue
        '
        Me.lblInViolationValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblInViolationValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInViolationValue.Enabled = False
        Me.lblInViolationValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInViolationValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInViolationValue.Location = New System.Drawing.Point(456, 256)
        Me.lblInViolationValue.Name = "lblInViolationValue"
        Me.lblInViolationValue.Size = New System.Drawing.Size(152, 17)
        Me.lblInViolationValue.TabIndex = 1079
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(368, 248)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 23)
        Me.Label5.TabIndex = 1078
        Me.Label5.Text = "In Violation:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Closure
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(972, 694)
        Me.Controls.Add(Me.tbCntrlClosure)
        Me.Controls.Add(Me.pnlTop)
        Me.Name = "Closure"
        Me.Text = "Closure"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.tbCntrlClosure.ResumeLayout(False)
        Me.tbPageOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerBottom.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
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
        Me.tbCtrlFacClosureEvts.ResumeLayout(False)
        Me.tbPageFacClosure.ResumeLayout(False)
        CType(Me.dgClosureFacilityDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityClosureButton.ResumeLayout(False)
        Me.tbPageFacilityDocuments.ResumeLayout(False)
        Me.pnl_FacilityDetail.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageClosures.ResumeLayout(False)
        Me.pnlClosureMain.ResumeLayout(False)
        Me.pnlClosuresDetails.ResumeLayout(False)
        Me.pnlContactDetails.ResumeLayout(False)
        Me.pnlClosureContactButtons.ResumeLayout(False)
        Me.pnlClosureContactsContainer.ResumeLayout(False)
        CType(Me.ugClosureContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlClosureContactHeader.ResumeLayout(False)
        Me.pnlContacts.ResumeLayout(False)
        Me.pnlClosureReportDetails.ResumeLayout(False)
        Me.pnlClosureReport.ResumeLayout(False)
        Me.pnlChecklistDetails.ResumeLayout(False)
        Me.pnlChecklist.ResumeLayout(False)
        Me.pnlAnalysisDetails.ResumeLayout(False)
        CType(Me.dGridAnalysis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAnalysis.ResumeLayout(False)
        Me.PnlTanksPipesDetails.ResumeLayout(False)
        CType(Me.udPreviousSubstance, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugTankandPipes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTanksPipes.ResumeLayout(False)
        Me.PnlNoticeOfInterestDetails.ResumeLayout(False)
        Me.pnlNoticeOfInterest.ResumeLayout(False)
        Me.pnlLUSTEventBottom.ResumeLayout(False)
        Me.pnlNoticeOfInterestHeader.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.tbPageSummary.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Intialization"
    Private Sub ClosureVisible()
        Try
            pClosure = pOwn.Facilities.ClosureEvent
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub InitControls()
        Try
            UIUtilsGen.PopulateOwnerType(cmbOwnerType, pOwn)
            ' UIUtilsGen.PopulateOrgEntityType(Me.cmbOwnerOrgEntityCode, pOwn)
            If pOwn.ID <> 0 Then
                UIUtilsGen.PopulateFacilityType(Me.cmbFacilityType, pOwn.Facilities)
                UIUtilsGen.PopulateFacilityDatum(Me.cmbFacilityDatum, pOwn.Facilities)
                UIUtilsGen.PopulateFacilityMethod(Me.cmbFacilityMethod, pOwn.Facilities)
                UIUtilsGen.PopulateFacilityLocationType(Me.cmbFacilityLocationType, pOwn.Facilities)
            End If
            SetupAnalysisDataTable()
            'btnFacilitySave.Enabled = False
            'btnFacilityCancel.Enabled = False
        Catch ex As Exception
            Throw ex
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
    Private Sub Closure_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Closure")

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Closure_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Dim MyFrm As MusterContainer
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Closure")
            MyFrm = Me.MdiParent
            'If lblOwnerIDValue.Text <> String.Empty And lblFacilityIDValue.Text = String.Empty Then
            'MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, "9")
            'ElseIf lblOwnerIDValue.Text <> String.Empty And lblFacilityIDValue.Text <> String.Empty Then
            'MyFrm.FlagsChanged(Me.lblFacilityIDValue.Text, "6")
            'End If

            MusterContainer.pOwn = pOwn

            If lblOwnerIDValue.Text <> String.Empty Then ' And lblFacilityIDValue.Text = String.Empty Then
                pOwn.Retrieve(Me.lblOwnerIDValue.Text, "ALL")
            End If
            bolFrmActivated = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Event Handlers for Expanding/Collapsing the different Sections of the Lust Event form"
    Private Sub ShowHideControl(ByVal ObjControl As Control)
        Try
            If ObjControl.Visible Then
                ObjControl.Visible = False
            Else
                ObjControl.Visible = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub lblNoticeOfInterestDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblNoticeOfInterestDisplay.Click
        Try
            If lblNoticeOfInterestDisplay.Text = "+" Then
                lblNoticeOfInterestDisplay.Text = "-"
            Else
                lblNoticeOfInterestDisplay.Text = "+"
            End If
            ShowHideControl(PnlNoticeOfInterestDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblNoticeOfInterestHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblNoticeOfInterestHead.Click
        Try
            lblNoticeOfInterestDisplay_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblTanksPipesDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTanksPipesDisplay.Click
        Try
            If lblTanksPipesDisplay.Text = "+" Then
                lblTanksPipesDisplay.Text = "-"
            Else
                lblTanksPipesDisplay.Text = "+"
            End If
            ShowHideControl(PnlTanksPipesDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblTanksPipesHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTanksPipesHead.Click
        Try
            lblTanksPipesDisplay_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblAnalysisDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblAnalysisDisplay.Click
        Try
            If lblAnalysisDisplay.Text = "+" Then
                lblAnalysisDisplay.Text = "-"
            Else
                lblAnalysisDisplay.Text = "+"
            End If
            ShowHideControl(pnlAnalysisDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblAnalysisHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblAnalysisHead.Click
        Try
            lblAnalysisDisplay_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblChecklistDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblChecklistDisplay.Click
        Try
            If pClosure.ClosureType = 444 Or pClosure.ClosureType = 445 Then
                If lblChecklistDisplay.Text = "+" Then
                    lblChecklistDisplay.Text = "-"
                Else
                    lblChecklistDisplay.Text = "+"
                End If
                ShowHideControl(pnlChecklistDetails)
            Else
                lblChecklistDisplay.Text = "+"
                pnlChecklistDetails.Visible = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblChecklistHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblChecklistHead.Click
        Try
            lblChecklistDisplay_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblClosureReportDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblClosureReportDisplay.Click
        Try
            If lblClosureReportDisplay.Text = "+" Then
                lblClosureReportDisplay.Text = "-"
            Else
                lblClosureReportDisplay.Text = "+"
            End If
            ShowHideControl(pnlClosureReportDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblClosureReportHead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblClosureReportHead.Click
        Try
            lblClosureReportDisplay_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblContactsDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblContactsDisplay.Click
        Try
            If lblContactsDisplay.Text = "+" Then
                lblContactsDisplay.Text = "-"
            Else
                lblContactsDisplay.Text = "+"
            End If
            ShowHideControl(pnlContactDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lblContactsHead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblContactsHead.Click
        Try
            lblContactsDisplay_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Tab Operations"
    Private Sub tbCntrlClosure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCntrlClosure.Click
        Dim MyFrm As MusterContainer

        Try
            If strActivetbPage = "TBPAGECLOSURES" Then
                'strActivetbPage = String.Empty
                Exit Sub
            End If
            MyFrm = Me.MdiParent
            MyFrm.mnuCancelAllOverDue.Enabled = False

            Select Case tbCntrlClosure.SelectedTab.Name.ToUpper
                Case "TBPAGEFACILITYDETAIL"
                    nCurrentEventID = -1
                    If Me.ugFacilityList.Rows.Count <> 0 Then
                        If Me.lblFacilityIDValue.Text = String.Empty Then
                            PopulateFacilityInfo(Integer.Parse(ugFacilityList.ActiveRow.Cells("FACILITYID").Text))
                        Else
                            PopulateFacilityInfo(Integer.Parse(Me.lblFacilityIDValue.Text))
                        End If
                        Me.Tag = Me.lblFacilityIDValue.Text
                        Me.lblFacilityIDValue.Focus()
                        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Closure")
                    End If
                    If ugFacilityList.Rows.Count <= 0 And Me.lblOwnerIDValue.Text <> String.Empty Then
                        Dim msgResult As MsgBoxResult
                        msgResult = MsgBox("No facilities found for owner" + lblOwnerIDValue.Text)
                        Exit Select
                    End If

                    Me.Text = "Closure - Facility Detail - (" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"
                Case "TBPAGECLOSURES"

                    If Me.dgClosureFacilityDetails.Rows.Count <> 0 Then
                        MyFrm.mnuCancelAllOverDue.Enabled = True
                        'Else
                        '    MyFrm.mnuCancelAllOverDue.Enabled = False
                    End If
                    If bolLoading Then Exit Sub
                    If Me.dgClosureFacilityDetails.Rows.Count <> 0 Then
                        If Integer.Parse(Me.lblClosureIDValue.Text) > 0 Then
                            Me.LoadClosureData(Integer.Parse(dgClosureFacilityDetails.ActiveRow.Cells("NOI ID").Text))
                        Else
                            LoadClosureData(Integer.Parse(dgClosureFacilityDetails.Rows(0).Cells("NOI ID").Text))
                        End If
                    End If
                    If dgClosureFacilityDetails.Rows.Count <= 0 And Me.lblFacilityIDValue.Text <> String.Empty Then 'Me.lblClosureIDValue.Text <> String.Empty Then
                        MsgBox("No Closures found for Facility" + lblFacilityIDValue.Text)
                        Me.tbCntrlClosure.SelectedTab = Me.tbPageFacilityDetail
                        tbCntrlClosure_Click(sender, e)
                        Exit Select
                    ElseIf Me.lblFacilityIDValue.Text = String.Empty Then
                        tbCntrlClosure.SelectedTab = tbPageFacilityDetail
                        tbCntrlClosure_Click(sender, e)
                        Exit Select

                    End If
                    'Me.Text = "Closure - Closure Event - (" & IIf(Me.lblClosureIDValue.Text = String.Empty, "New", lblClosureIDValue.Text) & ")"
                    Me.Text = "Closure - Closure Event - Facility #" & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
                Case "TBPAGEOWNERDETAIL"
                    nCurrentEventID = -1
                    Me.Text = "Closure - Owner Detail (" & txtOwnerName.Text & ")"
                    'Me.PopulateOwnerInfo(pOwn.ID)
                    If lblOwnerIDValue.Text <> String.Empty Then
                        UIUtilsGen.PopulateOwnerFacilities(pOwn, Me, Integer.Parse(Me.lblOwnerIDValue.Text))
                    End If

                    If pOwn.ID > 0 And Not TabControl1.TabPages.Contains(tbPageOwnerContactList) Then
                        TabControl1.TabPages.Add(tbPageOwnerContactList)
                        lblOwnerContacts.Text = "Owner Contacts"
                    End If
                    'LoadContacts(ugOwnerContacts, pOwn.ID, 9)
                    TabControl1.SelectedTab = tbPageOwnerFacilities
                    MyFrm = Me.MdiParent
                Case "TBPAGEOWNERFACILITIES"
                    nCurrentEventID = -1
                Case "TBPAGESUMMARY"
                    nCurrentEventID = -1
                    Me.Text = "Closure - Owner Summary (" & txtOwnerName.Text & ")"
                    UIUtilsGen.PopulateOwnerSummary(pOwn, Me)
            End Select
            strActivetbPage = tbCntrlClosure.SelectedTab.Name.ToString.ToUpper
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub tbCntrlClosure_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlClosure.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Dim MyFrm As MusterContainer
        Dim nFacilityID As Integer
        Try

            MyFrm = Me.MdiParent
            Select Case tbCntrlClosure.SelectedTab.Name.ToUpper
                Case "TBPAGEOWNERDETAIL"
                    nCurrentEventID = -1
                    Me.Text = "Closure - Owner Detail (" & txtOwnerName.Text & ")"
                    Me.PopulateOwnerInfo(pOwn.ID)
                    strActivetbPage = tbCntrlClosure.SelectedTab.Name.ToString.ToUpper
                    If lblOwnerIDValue.Text <> String.Empty Then
                        MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, UIUtilsGen.EntityTypes.Owner, "Closure", Me.Text)
                    End If
                    TabControl1.SelectedTab = tbPageOwnerFacilities
                Case "TBPAGEFACILITYDETAIL"
                    nCurrentEventID = -1
                    Me.Text = "Closure - Facility Detail - (" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"
                    strActivetbPage = tbCntrlClosure.SelectedTab.Name.ToString.ToUpper
                    If lblFacilityIDValue.Text <> String.Empty Then
                        MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text)
                    End If
                    If ugFacilityList.Rows.Count > 0 And Not pOwn.Facilities.ID > 0 Then
                        If ugFacilityList.ActiveRow Is Nothing Then
                            ugFacilityList.ActiveRow = ugFacilityList.Rows(0)
                        End If
                        nFacilityID = ugFacilityList.ActiveRow.Cells("FacilityID").Value
                        Me.PopulateFacilityInfo(nFacilityID)
                    Else
                        Me.PopulateFacilityInfo(pOwn.Facilities.ID)
                    End If
                Case "TBPAGECLOSURES"
                    '                    Me.Text = "Closure - Closure Event - (" & IIf(Me.lblClosureIDValue.Text = String.Empty, "New", lblClosureIDValue.Text) & ")"
                    Me.Text = "Closure - Closure Event - Facility #" & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
                    If lblFacilityIDValue.Text <> String.Empty Then
                        If Not MyFrm Is Nothing Then
                            nCurrentEventID = pClosure.ID
                            MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text, pClosure.ID, UIUtilsGen.EntityTypes.ClosureEvent)
                        End If
                    End If
                Case Else
                    nCurrentEventID = -1
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Form Events"
    Private Sub Closure_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cntrl As Control
        Dim uiUts As New UIUtilsGen
        Try
            pnlOwnerName.Visible = False
            For Each cntrl In Me.Controls
                uiUts.ClearComboBox(cntrl)
            Next
            bolLoading = True

            bolLoading = False

           




        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            If pOwn.colIsDirty() Then
                Dim results As Long = MsgBoxResult.No

                If _container.DirtyIgnored = -1 AndAlso MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed") = MsgBoxResult.Yes Then
                    results = MsgBoxResult.Yes
                Else
                    results = _container.DirtyIgnored
                End If

                If results = MsgBoxResult.Yes Then
                    Dim success As Boolean = False
                    pOwn.ModifiedBy = MusterContainer.AppUser.ID
                    btnSaveClosure_Click(Me, New EventArgs)

                    success = pOwn.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        e.Cancel = True
                        Exit Sub
                    End If
                    If Not success Then

                        e.Cancel = True
                        bolValidateSuccess = True
                        bolDisplayErrmessage = True
                        nErrMessage = 0
                        Exit Sub
                    End If
                ElseIf results = MsgBoxResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                Else
                    'Reset this closure event as if nothing has happen
                    pOwn.Facilities.ClosureEvent.Reset()

                End If
            End If
            'if any other forms are using the owner, leave alone. else remove from collection
            UIUtilsGen.RemoveOwner(pOwn, Me)
            Dim MyFrm As MusterContainer = Me.MdiParent
            If Not MyFrm Is Nothing Then
                MyFrm.mnuCancelAllOverDue.Enabled = False
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region
    '#Region "State lookup values"
    'Private Function ArrayListToDataTable(ByVal ArrLst As ArrayList) As DataTable
    '    Dim i As Int16
    '    Dim dsTable As New DataTable
    '    Dim dsDR As DataRow
    '    Try
    '        dsTable.Columns.Add("Id")
    '        dsTable.Columns(0).DataType = System.Type.GetType("System.Int32")
    '        dsTable.Columns.Add("Type")
    '        dsTable.Columns(1).DataType = System.Type.GetType("System.String")
    '        For i = 0 To ArrLst.Count - 1
    '            Dim LP As InfoRepository.LookupProperty
    '            LP = ArrLst(i)
    '            dsDR = dsTable.NewRow
    '            dsDR.Item(0) = LP.Id
    '            dsDR.Item(1) = LP.Type
    '            dsTable.Rows.Add(dsDR)
    '        Next
    '        Return dsTable
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function
    '#End Region
#Region "Owner Operations"
#Region "UI Support Routines"
    Friend Sub NewOwner()
        'pOwn = New MUSTER.BusinessLogic.pOwner
        Try
            pOwn.Retrieve(0)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub SetupAddOwner()
        Try
            ClearOwnerForm()
            cmbOwnerType.Focus()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ClearOwnerForm()
        Try
            bolLoading = True
            UIUtilsGen.ClearFields(pnlOwnerDetail)
            rdOwnerOrg.Tag = False
            rdOwnerPerson.Tag = True
            chkCAPParticipant.Tag = True 'Person Mode is the default for Owner Name
            cmbOwnerType.SelectedIndex = -1
            txtOwnerAddress.Tag = 0
            txtOwnerName.Tag = 0
            lblNoOfFacilitiesValue.Text = "0"
            bolLoading = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetPersonaSaveCancel(ByVal bolstate As Boolean)
        Try
            Me.btnOwnerNameOK.Enabled = True
            Me.btnOwnerNameCancel.Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetOwnerSaveCancel(ByVal bolstate As Boolean)
        Try
            Me.btnSaveOwner.Enabled = bolstate
            Me.btnOwnerCancel.Enabled = bolstate
            If pOwn.ID > 0 Then
                If Not Me.ugFacilityList.Rows Is Nothing Then
                    If Me.ugFacilityList.Rows.Count = 0 Then
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub PopulateOwnerInfo(ByVal OwnerID As Integer)
        Dim MyFrm As MusterContainer
        Try
            UIUtilsGen.PopulateOwnerInfo(OwnerID, pOwn, Me)
            Me.EnableDisableOwnerControls()

            If Not TabControl1.TabPages.Contains(tbPageOwnerContactList) Then
                TabControl1.TabPages.Add(tbPageOwnerContactList)
                lblOwnerContacts.Text = "Owner Contacts"
            End If
            Select Case TabControl1.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If OwnerID > 0 Then
                        LoadContacts(ugOwnerContacts, OwnerID, UIUtilsGen.EntityTypes.Owner)
                    End If
                Case tbPageOwnerDocuments.Name
                    UCOwnerDocuments.LoadDocumentsGrid(OwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Closure)
            End Select
            'LoadContacts(ugOwnerContacts, OwnerID, 9)
            If bolFrmActivated Then
                MyFrm = Me.MdiParent
                'oEntity.GetEntity("Owner")
                If lblOwnerIDValue.Text <> String.Empty Then
                    MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, UIUtilsGen.EntityTypes.Owner, "Closure", Me.Text)
                End If
            End If
            CommentsMaintenance(, , True)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Friend Function OwnerCAEFacilities() As DataSet
        Try
            Return pOwn.GetFacilitiesCAESummaryTable
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Friend Sub EnableDisableOwnerControls()
        Try
            If pOwn.ID > 0 Then
                Me.btnOwnerComment.Enabled = True
                Me.btnOwnerFlag.Enabled = True
            Else
                Me.btnOwnerComment.Enabled = False
                Me.btnOwnerFlag.Enabled = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub txtOwnerName_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOwnerName.DoubleClick
        Try
            If Me.txtOwnerName.Text <> String.Empty Then
                UIUtilsGen.setupOwnername(Me, pOwn)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub ugFacilityList_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugFacilityList.DoubleClick
    '    Me.tbCntrlClosure.SelectedIndex = 1
    '    PopulateFacilityInfo(ugFacilityList.ActiveRow.Cells.Item("FacilityID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw))
    'End Sub
    Private Sub btnSaveOwner_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveOwner.Click
        Try
            pOwn.ModifiedBy = MusterContainer.AppUser.ID
            pOwn.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If bolValidateSuccess Then
                If lblOwnerIDValue.Text.Trim.Length > 0 Then
                    MsgBox("The Owner has been Successfully Modified to the System!")
                    PopulateOwnerInfo(CInt(lblOwnerIDValue.Text))
                End If

            Else
                bolValidateSuccess = True
                bolDisplayErrmessage = True
                nErrMessage = 0
                Exit Sub
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtOwnerAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.TextChanged
        If bolLoading Then Exit Sub
        Try
            If txtOwnerAddress.Tag > 0 Then
                pOwn.AddressId = Integer.Parse(Trim(txtOwnerAddress.Tag))
            Else
                pOwn.AddressId = 0
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mskTxtOwnerPhone_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtOwnerPhone.KeyUpEvent
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillStringObjectValues(pOwn.PhoneNumberOne, mskTxtOwnerPhone.FormattedText.Trim.ToString)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mskTxtOwnerPhone2_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtOwnerPhone2.KeyUpEvent
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillStringObjectValues(pOwn.PhoneNumberTwo, mskTxtOwnerPhone2.FormattedText.Trim.ToString)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mskTxtOwnerFax_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtOwnerFax.KeyUpEvent
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillStringObjectValues(pOwn.Fax, mskTxtOwnerFax.FormattedText.Trim.ToString)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbOwnerType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.OwnerType = UIUtilsGen.GetComboBoxValueInt(cmbOwnerType)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtOwnerEmail_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerEmail.TextChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.EmailAddress = Me.txtOwnerEmail.Text
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkOwnerAgencyInterest_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOwnerAgencyInterest.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.EnsiteAgencyInterestID = chkOwnerAgencyInterest.Checked
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerCancel.Click
        Try
            If Not pOwn Is Nothing Then
                pOwn.Reset()
                Me.ClearOwnerForm()
                If pOwn.ID > 0 Then
                    Me.PopulateOwnerInfo(pOwn.ID)
                Else
                    pOwn.BPersona.Clear()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
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
#Region "Owner Name and Address UI handlers"
    'Private Sub AddrForm_NewAddressData(ByVal MyAddress As String) Handles AddrForm.NewAddressData
    '    Try
    '        Select Case tbCntrlClosure.SelectedTab.Name.ToUpper
    '            Case "TBPAGEOWNERDETAIL"
    '                txtOwnerAddress.Text = MyAddress
    '            Case "TBPAGEFACILITYDETAIL"
    '                txtFacilityAddress.Text = MyAddress
    '        End Select
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub AddressForm_NewAddressID(ByVal MyAddressID As Integer) Handles AddressForm.NewAddressID
        Try
            Select Case tbCntrlClosure.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    txtOwnerAddress.Tag = MyAddressID
                    pOwn.AddressId = MyAddressID
                Case tbPageFacilityDetail.Name
                    txtFacilityAddress.Tag = MyAddressID
                    pOwn.Facilities.AddressID = MyAddressID
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub AddrForm_AddressDetails(ByVal addr1 As String, ByVal addr2 As String, ByVal cty As String, ByVal st As String, ByVal zip As String, ByVal fips As String) Handles AddrForm.AddressDetails
    '    Try
    '        Select Case tbCntrlClosure.SelectedTab.Name.ToUpper
    '            Case "TBPAGEOWNERDETAIL"
    '                pOwn.Addresses.AddressLine1 = addr1
    '                pOwn.Addresses.AddressLine2 = addr2
    '                pOwn.Addresses.City = cty
    '                pOwn.Addresses.State = st
    '                pOwn.Addresses.Zip = zip
    '                pOwn.Addresses.FIPSCode = fips
    '            Case "TBPAGEFACILITYDETAIL"
    '                pOwn.Facilities.FacilityAddresses.AddressLine1 = addr1
    '                pOwn.Facilities.FacilityAddresses.AddressLine2 = addr2
    '                pOwn.Facilities.FacilityAddresses.City = cty
    '                pOwn.Facilities.FacilityAddresses.State = st
    '                pOwn.Facilities.FacilityAddresses.Zip = zip
    '                pOwn.Facilities.FacilityAddresses.FIPSCode = fips
    '        End Select
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
  
    Private Sub txtOwnerAddress_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.DoubleClick
        Try

            Dim addressForm As Address
            Address.EditAddress(addressForm, pOwn.ID, pOwn.Addresses, "Owner", UIUtilsGen.ModuleID.Closure, txtOwnerAddress, UIUtilsGen.EntityTypes.Owner)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub txtOwnerAddress_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtOwnerAddress.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                txtOwnerAddress_DblClick(sender, New System.EventArgs)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtOwnerAddress_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.Enter
        Try
            If txtOwnerAddress.Text = String.Empty Then
                txtOwnerAddress_DblClick(sender, e)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub SwapOrgPersonDisplay()
        Try
            UIUtilsGen.SwapOrgPersonDisplay(Me)
            'If rdOwnerPerson.Checked Then
            '    pnlOwnerPerson.Location = New Point(pnlPersonOrganization.Location.X, pnlPersonOrganization.Location.Y + pnlPersonOrganization.Height)
            '    pnlOwnerName.Height = pnlPersonOrganization.Height + pnlOwnerPerson.Height + 5 + pnlOwnerNameButton.Height
            '    pnlOwnerNameButton.Location = New Point(pnlOwnerPerson.Location.X, pnlOwnerPerson.Location.Y + pnlOwnerPerson.Height)
            '    pnlOwnerOrg.Visible = False
            '    pnlOwnerPerson.Visible = True
            'Else
            '    pnlOwnerOrg.Location = New Point(pnlPersonOrganization.Location.X, pnlPersonOrganization.Location.Y + pnlPersonOrganization.Height)
            '    pnlOwnerName.Height = pnlPersonOrganization.Height + pnlOwnerOrg.Height + pnlOwnerNameButton.Height
            '    pnlOwnerNameButton.Location = New Point(pnlOwnerOrg.Location.X, pnlOwnerOrg.Location.Y + pnlOwnerOrg.Height)
            '    pnlOwnerOrg.Visible = True
            '    pnlOwnerPerson.Visible = False
            '    If bolNewPersona = True Then
            '        bolLoading = True
            '        Me.cmbOwnerOrgEntityCode.SelectedIndex = -1
            '        bolNewPersona = False
            '        bolLoading = False
            '    End If
            'End If
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Sub
    Private Sub rdOwnerPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdOwnerPerson.Click
        Try
            UIUtilsGen.rdOwnerPersonClick(Me, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
        'Dim msgResult As MsgBoxResult
        'If pOwn.BPersona.OrgID <> 0 Or pOwn.BPersona.IsDirty Then
        '    If rdOwnerOrg.Tag = True And rdOwnerPerson.Tag = False Then
        '        msgResult = MsgBox(" Do you want to change from Organization to Person", MsgBoxStyle.YesNo, "Persona")
        '        If msgResult = MsgBoxResult.Yes Then
        '            ClearPersona()
        '            'ClearBPersonaOrganization()
        '            pOwn.BPersona.Clear()
        '            bolNewPersona = True
        '            SwapOrgPersonDisplay()
        '            rdOwnerOrg.Tag = False
        '            rdOwnerPerson.Tag = True
        '        Else
        '            rdOwnerOrg.Checked = True
        '        End If
        '    End If
        'Else
        '    SwapOrgPersonDisplay()
        '    '    ClearPersona()
        'End If

    End Sub
    Private Sub rdOwnerOrg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdOwnerOrg.Click
        Try
            UIUtilsGen.rdOwnerOrgClick(Me, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
        'Dim msgResult As MsgBoxResult
        'If pOwn.BPersona.PersonId <> 0 Or pOwn.BPersona.IsDirty Then
        '    If rdOwnerOrg.Tag = False And rdOwnerPerson.Tag = True Then
        '        msgResult = MsgBox(" Do you want to change from Person to Organization", MsgBoxStyle.YesNo, "Persona")
        '        If msgResult = MsgBoxResult.Yes Then
        '            ClearPersona()
        '            pOwn.BPersona.Clear()
        '            bolNewPersona = True
        '            SwapOrgPersonDisplay()
        '            rdOwnerOrg.Tag = True
        '            rdOwnerPerson.Tag = False
        '        Else
        '            rdOwnerPerson.Checked = True
        '        End If
        '    End If
        'Else
        '    SwapOrgPersonDisplay()
        '    bolLoading = True
        '    Me.cmbOwnerOrgEntityCode.SelectedIndex = -1
        '    bolLoading = False
        'End If
    End Sub
    Private Sub txtOwnerName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerName.Enter
        Try
            UIUtilsGen.setupOwnername(Me, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub pnlPersonOrganization_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlPersonOrganization.LostFocus
        Try
            cmbOwnerNameTitle.Focus()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbOwnerNameTitle_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerNameTitle.SelectedIndexChanged
        If bolLoading Then Exit Sub
        'If Me.cmbOwnerNameTitle.SelectedIndex <> -1 Then
        '    pOwn.BPersona.Title = Me.cmbOwnerNameTitle.Text
        'Else
        '    pOwn.BPersona.Title = String.Empty
        'End If
        Try
            UIUtilsGen.FillPersona(cmbOwnerNameTitle, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub txtOwnerFirstName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerFirstName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerFirstName, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
        'If Me.txtOwnerFirstName.Text <> String.Empty Then
        '    pOwn.BPersona.FirstName = Me.txtOwnerFirstName.Text
        'Else
        '    pOwn.BPersona.FirstName = String.Empty
        'End If
    End Sub
    Private Sub txtOwnerLastName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerLastName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerLastName, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
        'If Me.txtOwnerLastName.Text <> String.Empty Then
        '    pOwn.BPersona.LastName = Me.txtOwnerLastName.Text
        'Else
        '    pOwn.BPersona.LastName = String.Empty
        'End If
    End Sub
    Private Sub txtOwnerMiddleName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerMiddleName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerMiddleName, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
        'If Me.txtOwnerMiddleName.Text <> String.Empty Then
        '    pOwn.BPersona.MiddleName = Me.txtOwnerMiddleName.Text
        'Else
        '    pOwn.BPersona.MiddleName = String.Empty
        'End If
    End Sub
    Private Sub cmbOwnerNameSuffix_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerNameSuffix.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(cmbOwnerNameSuffix, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
        'If cmbOwnerNameSuffix.SelectedIndex <> -1 Then
        '    pOwn.BPersona.Suffix = Me.cmbOwnerNameSuffix.Text
        'Else
        '    pOwn.BPersona.Suffix = String.Empty
        'End If
    End Sub
    Private Sub txtOwnerOrgName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerOrgName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerOrgName, pOwn)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
        'If Me.txtOwnerOrgName.Text <> String.Empty Then
        '    pOwn.BPersona.Company = Me.txtOwnerOrgName.Text
        'Else
        '    pOwn.BPersona.Company = String.Empty
        'End If
    End Sub
    'Private Sub cmbOwnerOrgEntityCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   If bolLoading Then Exit Sub
    '  Try
    '     UIUtilsGen.FillPersona(cmbOwnerOrgEntityCode, pOwn)
    ' Catch ex As Exception
    '     Dim MyErr As New ErrorReport(ex)
    '     MyErr.ShowDialog()
    ' End Try
    '''If Me.cmbOwnerOrgEntityCode.SelectedValue > 0 Then
    '''    pOwn.BPersona.Org_Entity_Code = Me.cmbOwnerOrgEntityCode.SelectedValue
    '''Else
    '''    pOwn.BPersona.Org_Entity_Code = 0

    '''End If
    'End Sub

    Private Sub btnOwnerNameOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerNameOK.Click
        Try

            Dim success As Boolean = False
            If pOwn.BPersona.PersonId > 0 Or pOwn.BPersona.Org_Entity_Code > 0 Then
                pOwn.BPersona.ModifiedBy = MusterContainer.AppUser.ID
            Else

            End If
            success = pOwn.BPersona.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If success Then
                UIUtilsGen.SetOwnerName(Me)
                txtOwnerName.Tag = Nothing
                pnlOwnerName.Hide()
                lblOwnerAddress.Focus()
            Else
                bolValidateSuccess = True
                bolDisplayErrmessage = True
                nErrMessage = 0
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
            Exit Sub
        End Try
        txtOwnerAddress.Focus()

    End Sub
    'Private Sub SetOwnerName()
    '    Dim strOwnerName As String = String.Empty
    '    If rdOwnerOrg.Checked Then
    '        strOwnerName = txtOwnerOrgName.Text
    '    Else
    '        strOwnerName = IIf(cmbOwnerNameTitle.Text.Trim.Length > 0, cmbOwnerNameTitle.Text + " ", "") + txtOwnerFirstName.Text + " " + IIf(txtOwnerMiddleName.Text.Trim.Length > 0, txtOwnerMiddleName.Text + " ", "") + txtOwnerLastName.Text + IIf(cmbOwnerNameSuffix.Text.Trim.Length > 0, " " + cmbOwnerNameSuffix.Text, "")
    '    End If
    '    txtOwnerName.Text = strOwnerName

    'End Sub
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
            Me.SetOwnerSaveCancel(True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub CheckUncheckPersonaOrg(ByVal bolPerson As Boolean, ByVal bolOrg As Boolean)
    '    rdOwnerOrg.Checked = bolOrg
    '    rdOwnerPerson.Checked = bolPerson
    '    rdOwnerOrg.Tag = bolOrg
    '    rdOwnerPerson.Tag = bolPerson
    '    SwapOrgPersonDisplay()
    'End Sub

    'Private Sub ClearPersona()

    '    If Me.rdOwnerPerson.Checked = False And Me.rdOwnerOrg.Checked = False Then
    '        Me.rdOwnerPerson.Checked = True
    '    End If
    '    bolLoading = True
    '    txtOwnerOrgName.Text = String.Empty
    '    cmbOwnerNameTitle.SelectedIndex = -1
    '    cmbOwnerNameSuffix.SelectedIndex = -1
    '    txtOwnerFirstName.Text = String.Empty
    '    txtOwnerLastName.Text = String.Empty
    '    txtOwnerMiddleName.Text = String.Empty
    '    cmbOwnerOrgEntityCode.SelectedIndex = -1
    '    cmbOwnerOrgEntityCode.SelectedIndex = -1
    '    bolLoading = False
    'End Sub

    'Private Function ResetOwnerName() As String
    '    Dim oPersonaInfo As MUSTER.Info.PersonaInfo
    '    Dim strOwnerName As String = String.Empty
    '    Try
    '        If pOwn.PersonID = 0 Then
    '            Me.CheckUncheckPersonaOrg(False, True)
    '            oPersonaInfo = pOwn.Organization()
    '            txtOwnerOrgName.Text = IIf(IsNothing(pOwn.BPersona.Company), String.Empty, CStr(pOwn.BPersona.Company))
    '            UIUtilsGen.ValidateComboBoxItemByValue(cmbOwnerOrgEntityCode, pOwn.BPersona.Org_Entity_Code)
    '            strOwnerName = Me.txtOwnerOrgName.Text
    '        Else
    '            Me.CheckUncheckPersonaOrg(True, False)
    '            oPersonaInfo = pOwn.Persona()
    '            cmbOwnerNameTitle.Text = IIf(IsNothing(Trim(pOwn.BPersona.Title)), String.Empty, CStr(Trim(pOwn.BPersona.Title)))
    '            txtOwnerFirstName.Text = IIf(IsNothing(pOwn.BPersona.FirstName), String.Empty, CStr(pOwn.BPersona.FirstName))
    '            txtOwnerLastName.Text = IIf(IsNothing(pOwn.BPersona.LastName), String.Empty, CStr(pOwn.BPersona.LastName))
    '            cmbOwnerNameSuffix.Text = IIf(pOwn.BPersona.Suffix = String.Empty, String.Empty, CStr(Trim(pOwn.BPersona.Suffix)))
    '            txtOwnerMiddleName.Text = IIf(IsNothing(pOwn.BPersona.MiddleName), String.Empty, CStr(pOwn.BPersona.MiddleName))
    '            strOwnerName = IIf(cmbOwnerNameTitle.Text.Trim.Length > 0, cmbOwnerNameTitle.Text.ToString() + " ", "") + txtOwnerFirstName.Text.ToString() + " " + IIf(txtOwnerMiddleName.Text.Trim.Length > 0, txtOwnerMiddleName.Text.ToString() + " ", "") + Me.txtOwnerLastName.Text.ToString() + IIf(cmbOwnerNameSuffix.Text.Trim.Length > 0, " " + cmbOwnerNameSuffix.Text.ToString(), "")
    '        End If
    '        Return strOwnerName
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function
    Private Sub btnOwnerNameClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerNameClose.Click
        Dim bolState As Boolean
        Try
            bolState = Me.CheckObjectState(pOwn.BPersona)
            If Not bolState And bolValidateSuccess Then
                pnlOwnerName.Hide()
            End If
            bolValidateSuccess = True
            bolDisplayErrmessage = True
            nErrMessage = 0
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerNameClose_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerNameClose.LostFocus
        Try
            If Me.rdOwnerOrg.Checked = True Then
                Me.txtOwnerOrgName.Focus()
            Else
                Me.cmbOwnerNameTitle.Focus()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
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
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region
#Region "Facility Operations"
#Region "Facility Form Setup"
    Friend Sub SetupAddFacilityForm()
        Try
            If Not tbCntrlClosure.TabPages.Contains(tbPageFacilityDetail) Then
                tbCntrlClosure.TabPages.Add(tbPageFacilityDetail)
            End If
            tbCntrlClosure.SelectedTab = tbPageFacilityDetail
            btnFacilitySave.Enabled = False
            ClearFacilityForm()
            txtFacilityName.Focus()
            If Me.ugFacilityList.Rows.Count > 0 Then
                Me.lnkLblNextFac.Enabled = True
                Me.lnkLblPrevFacility.Enabled = True
            Else
                Me.lnkLblNextFac.Enabled = False
                Me.lnkLblPrevFacility.Enabled = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ClearFacilityForm()
        Try
            UIUtilsGen.ClearFacilityForm(Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "UI Support Routines"
    Private Sub SetFacilitySaveCancel(ByVal bolState As Boolean)
        Try
            Me.btnFacilitySave.Enabled = bolState
            btnFacilityCancel.Enabled = bolState
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub PopulateFacilityInfo(Optional ByVal FacilityID As Integer = 0)
        Dim MyFrm As MusterContainer
        Try
            UIUtilsGen.PopulateFacilityInfo(Me, pOwn.OwnerInfo, pOwn.Facilities, FacilityID)
            nFacilityID = FacilityID
            Dim ds As DataSet = Me.OwnerCAEFacilities

            If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso FacilityID > 0 Then
                With ds.Tables(0)
                    For Each rec As DataRow In .Select(String.Format("FACILITYID = {0}", FacilityID))
                        lblInViolationValue.Text = rec.Item("In Violation")
                    Next

                End With
            End If

            If lblInViolationValue.Text.ToUpper = "YES" Then
                lblInViolationValue.BackColor = Color.Red
            Else
                lblInViolationValue.BackColor = System.Drawing.SystemColors.Control
            End If



            EnableDiableFacilityControls()
            ClosureVisible()
            getClosureEventsGrid()
            'pClosure.Retrieve(pOwn.Facility, , pOwn.Facilities.ID)
            MyFrm = Me.MdiParent
            If Not MyFrm Is Nothing Then
                If Me.dgClosureFacilityDetails.Rows.Count > 0 Then
                    MyFrm.mnuCancelAllOverDue.Enabled = True
                Else
                    MyFrm.mnuCancelAllOverDue.Enabled = False
                End If
            End If
            If FacilityID > 0 And Not tbCtrlFacClosureEvts.TabPages.Contains(tbPageOwnerContactList) Then
                tbCtrlFacClosureEvts.TabPages.Add(tbPageOwnerContactList)
                lblOwnerContacts.Text = "Facility Contacts"
            End If
            Select Case tbCtrlFacClosureEvts.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If Me.lblFacilityIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility)
                    End If
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Closure)
            End Select
            'LoadContacts(ugOwnerContacts, FacilityID, 6)
            If bolFrmActivated Then
                MyFrm = Me.MdiParent
                'oEntity.GetEntity("Facility")
                MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text)
            End If

            tbCtrlFacClosureEvts.SelectedTab = tbPageFacClosure
            CommentsMaintenance(, , True)

            'Added by Hua Cao 10/08/2008 Issue #: [ UST-3204] Summary: Need a date field labeled " TOS Assessment Date:" added to several modules
            ' Retreive info from tblReg_AssessDate

            Me.dtPickAssess.Enabled = True
            Dim dtNullDate As Date = CDate("01/01/0001")
            Dim sqlStr As String
            Dim dtReturn As DataTable
            sqlStr = "tblReg_AssessDate where FacilityId = " + lblFacilityIDValue.Text
            dtReturn = pClosure.GetDataTable(sqlStr)
            If dtReturn.Rows.Count > 0 Then
                If Not dtReturn.Rows(0).Item("AssessDate") Is System.DBNull.Value Then
                    Me.dtPickAssess.Format = DateTimePickerFormat.Short
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
            Else
                Me.dtPickAssess.Checked = False
                UIUtilsGen.SetDatePickerValue(dtPickAssess, dtNullDate)
            End If
            dtReturn.Clear()

            'Added by Hua Cao 08/07/2009 for facility fee balance
            sqlStr = "vFees_FacilitySummaryGrid where Facility_Id = " + lblFacilityIDValue.Text
            dtReturn = pClosure.GetDataTable(sqlStr)
            If dtReturn.Rows.Count > 0 Then
                txtFeeBalance.Text = dtReturn.Rows(0).Item("TodateBalance")
            Else
                txtFeeBalance.Text = ""
            End If
            dtReturn.Clear()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub EnableDiableFacilityControls()
        Try
            'If pOwn.Facilities.ID > 0 Then
            If Me.lblFacilityIDValue.Text <> String.Empty Then
                If pOwn.Facilities.ID > 0 Then
                    Me.btnFacComments.Enabled = True
                    Me.btnFacFlags.Enabled = True
                Else
                    Me.btnFacComments.Enabled = False
                    Me.btnFacFlags.Enabled = False
                End If

            Else
                Me.btnFacComments.Enabled = False
                Me.btnFacFlags.Enabled = False
            End If

            Me.lnkLblNextFac.Enabled = True
            Me.lnkLblPrevFacility.Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub getClosureEventsGrid()
        Dim drRow As DataRow
        Dim rowcount As Integer = 0
        Dim str As String = String.Empty
        Try
            strClosureEventIdTags = String.Empty
            dgClosureFacilityDetails.DataSource = pOwn.Facilities.ClosureEventDataSet()
            For Each drRow In pOwn.Facilities.ClosureEventDataSet.Tables(0).Rows
                If rowcount < pOwn.Facilities.ClosureEventDataSet.Tables(0).Rows.Count - 1 Then
                    str = ","
                Else
                    str = ""
                End If
                strClosureEventIdTags += drRow("NOI ID").ToString + str
                rowcount += 1
            Next
            dgClosureFacilityDetails.DataSource = pOwn.Facilities.ClosureEventDataSet()
            Me.dgClosureFacilityDetails.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            Me.dgClosureFacilityDetails.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False

            Me.dgClosureFacilityDetails.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            Me.dgClosureFacilityDetails.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("NOI ID").Hidden = True
            Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("FACILITY_ID").Hidden = True

            'Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("Closure Status").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("Received").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("CR Received").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("Type").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("AAL").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("Sent to Tech").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("NFA By Closure").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'Me.dgClosureFacilityDetails.DisplayLayout.Bands(0).Columns("NFA By Tech").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            lblNoOfClosuresValue.Text = dgClosureFacilityDetails.Rows.Count
            lblClosureCountValue.Text = " of " + lblNoOfClosuresValue.Text
            lblClosureCountValue.Left = lblClosureIDValue.Left + lblClosureIDValue.Width

        Catch ex As Exception
            Throw ex
        End Try
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
#End Region
#Region "UI Control Events"
    Private Sub chkCAPCandidate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCAPCandidate.CheckedChanged
        Try
            pOwn.Facilities.CAPCandidate = chkCAPCandidate.Checked
            If pOwn.Facilities.CAPCandidate And Not pOwn.Facilities.CAPCandidateOriginal Then
                pOwn.Facilities.GetCapStatus(pOwn.Facilities.ID)
                UIUtilsGen.PopulateFacilityCapInfo(Me, pOwn.Facilities)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lnkLblNextFac_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextFac.LinkClicked
        Try
            If IsNumeric(lblFacilityIDValue.Text) Then
                PopulateFacilityInfo(GetPrevNextFacility(lblFacilityIDValue.Text, True))
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Closure")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lnkLblPrevFacility_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevFacility.LinkClicked
        Try
            If IsNumeric(lblFacilityIDValue.Text) Then
                PopulateFacilityInfo(GetPrevNextFacility(lblFacilityIDValue.Text, False))
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Closurel")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFacilitySave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilitySave.Click
        Try

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
                cmdSQLCommand.CommandText = "update tblReg_AssessDate set AssessDate = " + IIf(dtPickAssess.Checked, "'" + dtPickAssess.Value.ToShortDateString.ToString + "'", "NULL") + " where FacilityID = " + lblFacilityIDValue.Text
            Else
                cmdSQLCommand.CommandText = "insert into tblReg_AssessDate values(" + lblFacilityIDValue.Text + "," + IIf(dtPickAssess.Checked, "'" + dtPickAssess.Value.ToShortDateString.ToString + "'", "NULL") + ")"
            End If
          
            aReader.Close()
            cmdSQLCommand.ExecuteNonQuery()
            conSQLConnection.Close()
            cmdSQLCommand.Dispose()
            conSQLConnection.Dispose()

            pOwn.Facilities.ModifiedBy = MusterContainer.AppUser.ID
            pOwn.Facilities.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)

            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If bolValidateSuccess Then
                If Me.lblFacilityIDValue.Text.Length > 0 Then
                    MsgBox("The Facility has been Successfully Modified to the System!")
                End If
                PopulateFacilityInfo(Integer.Parse(Me.lblFacilityIDValue.Text))
            Else
                bolValidateSuccess = True
                bolDisplayErrmessage = True
                nErrMessage = 0
                Exit Sub
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFacilityCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityCancel.Click
        Try

            If Not pOwn.Facilities Is Nothing Then
                pOwn.Facilities.Reset()
                Me.ClearFacilityForm()
                If pOwn.Facilities.ID > 0 Then
                    Me.PopulateFacilityInfo(pOwn.Facilities.ID)
                End If
                Me.btnFacilityCancel.Enabled = False
                Me.btnFacilitySave.Enabled = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickFacilityRecvd_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickFacilityRecvd.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickFacilityRecvd)
            UIUtilsGen.FillDateobjectValues(pOwn.Facilities.DateReceived, dtPickFacilityRecvd.Text)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugFacilityList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFacilityList.DoubleClick
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
            tbCntrlClosure.SelectedTab = Me.tbPageFacilityDetail
            'Me.tbCntrlClosure.SelectedIndex = 1
            PopulateFacilityInfo(Integer.Parse(ugFacilityList.ActiveRow.Cells.Item("FacilityID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)))
            Me.Tag = Me.lblFacilityIDValue.Text
            'pClosure.Retrieve(pOwn.Facility, , pOwn.Facilities.ID)
            If (ugFacilityList.ActiveRow.Cells.Item("Total").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)) = "0" Then
                Me.btnAddClosure.Enabled = False
            Else
                Me.btnAddClosure.Enabled = True
            End If
            MyFrm = Me.MdiParent
            MyFrm.FlagsChanged(Me.lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub
    Private Sub txtFacilityName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityName.TextChanged
        Dim dtDateTransferred As Date
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.Name = Me.txtFacilityName.Text
            If Me.lblDateTransfered.Text <> String.Empty Then
                dtDateTransferred = lblDateTransfered.Text
                If Date.Compare(pOwn.Facilities.DateTransferred, dtDateTransferred) <> 0 Then
                    pOwn.Facilities.DateTransferred = dtDateTransferred 'CType(Trim(lblDateTransfered.Text), Date)
                End If
            End If
            If Me.txtFacilityName.Text <> String.Empty Then
                pOwn.Facilities.OwnerID = Integer.Parse(lblOwnerIDValue.Text.Trim)
            Else
                pOwn.Facilities.OwnerID = 0
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mskTxtFacilityPhone_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtFacilityPhone.KeyUpEvent
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillStringObjectValues(pOwn.Facilities.Phone, mskTxtFacilityPhone.FormattedText.ToString)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub mskTxtFacilityFax_KeyUpEvent(ByVal sender As System.Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtFacilityFax.KeyUpEvent
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillStringObjectValues(pOwn.Facilities.Fax, mskTxtFacilityFax.FormattedText.ToString)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFacilityLongDegree_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLongDegree.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillSingleObjectValues(pOwn.Facilities.LongitudeDegree, txtFacilityLongDegree.Text.Trim)
            UIUtilsGen.Check_If_Datum_Enable(Me)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFacilityLongMin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLongMin.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillSingleObjectValues(pOwn.Facilities.LongitudeMinutes, txtFacilityLongMin.Text.Trim)
            UIUtilsGen.Check_If_Datum_Enable(Me)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFacilityLongSec_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLongSec.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillDoubleObjectValues(pOwn.Facilities.LongitudeSeconds, txtFacilityLongSec.Text.Trim)
            UIUtilsGen.Check_If_Datum_Enable(Me)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFacilityLatDegree_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLatDegree.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillSingleObjectValues(pOwn.Facilities.LatitudeDegree, txtFacilityLatDegree.Text.Trim)
            UIUtilsGen.Check_If_Datum_Enable(Me)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFacilityLatMin_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLatMin.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillSingleObjectValues(pOwn.Facilities.LatitudeMinutes, txtFacilityLatMin.Text.Trim)
            UIUtilsGen.Check_If_Datum_Enable(Me)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFacilityLatSec_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityLatSec.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillDoubleObjectValues(pOwn.Facilities.LatitudeSeconds, txtFacilityLatSec.Text.Trim)
            UIUtilsGen.Check_If_Datum_Enable(Me)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFuelBrand_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFuelBrand.TextChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.FuelBrand = txtFuelBrand.Text.Trim.ToString
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbFacilityType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFacilityType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pOwn.Facilities.FacilityType = UIUtilsGen.GetComboBoxValueInt(cmbFacilityType)
            txtFacilitySIC.Text = cmbFacilityType.Text
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFacilityAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.TextChanged
        If bolLoading Then Exit Sub
        Try
            If txtFacilityAddress.Tag > 0 Then
                pOwn.Facilities.AddressID = Integer.Parse(Trim(txtFacilityAddress.Tag))
            Else
                pOwn.Facilities.AddressID = 0
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub ugFacilityList_AfterSortChange(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugFacilityList.AfterSortChange
        Try
            If ugFacilityList.Rows.Count > 0 Then
                ugFacilityList.ActiveRow = ugFacilityList.Rows(0)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtFacilityPowerOff_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtFacilityPowerOff.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtFacilityPowerOff)
            UIUtilsGen.FillDateobjectValues(pOwn.Facilities.DatePowerOff, dtFacilityPowerOff.Text)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkSignatureofNF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSignatureofNF.CheckedChanged
        Try
            If chkSignatureofNF.Checked Then
                txtDueByNF.Text = String.Empty
                '                SignatureFlag = False
            Else
                txtDueByNF.Text = "Due"
            End If
            pOwn.Facilities.SignatureOnNF = chkSignatureofNF.Checked
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub dtPickUpcomingInstallDateValue_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickUpcomingInstallDateValue.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickUpcomingInstallDateValue)
            'If Not pOwn.Facilities.UpcomingInstallationDate = "#12:00:00AM#" And dtPickUpcomingInstallDateValue.Text = "__/__/____" Then
            '    'bolUpcomingInstallation = False
            'End If
            UIUtilsGen.FillDateobjectValues(pOwn.Facilities.UpcomingInstallationDate, dtPickUpcomingInstallDateValue.Text)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkUpcomingInstall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUpcomingInstall.Click
        Dim result As MsgBoxResult
        Try
            If chkUpcomingInstall.Checked = False Then
                result = MessageBox.Show("Are you Sure you want to Clear this Date?", "Upcoming Install", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then
                    pOwn.Facilities.UpcomingInstallation = False
                    Exit Sub
                End If
                'bolUpcomingInstallation = False
                UIUtilsGen.CreateEmptyFormatDatePicker(dtPickUpcomingInstallDateValue)
                dtPickUpcomingInstallDateValue.Text = String.Empty
                dtPickUpcomingInstallDateValue.Enabled = False
                pOwn.Facilities.UpcomingInstallation = False
            ElseIf chkUpcomingInstall.Checked = True Then
                dtPickUpcomingInstallDateValue.Enabled = True
                pOwn.Facilities.UpcomingInstallation = True
                chkSignatureofNF.Checked = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dgClosureFacilityDetails_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgClosureFacilityDetails.DoubleClick
        Dim oGrid As Infragistics.Win.UltraWinGrid.UltraGrid
        Try
            oGrid = CType(sender, Infragistics.Win.UltraWinGrid.UltraGrid)

            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If

            GridClosureSelected(oGrid)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
        End Try
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
#Region "Facility Address UI Handler"
    Private Sub txtFacilityAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.DoubleClick
        Try

            Dim addressForm As Address
            If (pOwn.Facilities.ID = 0 Or pOwn.Facilities.ID <> nFacilityID) And nFacilityID > 0 Then
                UIUtilsGen.PopulateFacilityInfo(Me, pOwn.OwnerInfo, pOwn.Facilities, nFacilityID)

            End If

            Address.EditAddress(addressForm, pOwn.Facilities.ID, pOwn.Facilities.FacilityAddresses, "Facility", UIUtilsGen.ModuleID.Closure, txtFacilityAddress, UIUtilsGen.EntityTypes.Facility, True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub txtFacilityAddress_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFacilityAddress.KeyUp
        Try
            If e.KeyCode = Keys.Enter Then
                txtFacilityAddress_DoubleClick(sender, New System.EventArgs)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtFacilityAddress_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.Enter
        Try
            If txtFacilityAddress.Text = String.Empty Then
                txtFacilityAddress_DoubleClick(sender, e)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region
#Region "Misc Operations"
    Private Function CheckObjectState(ByRef obj As Object, Optional ByVal bolColDirty As Boolean = False) As Boolean ', Optional ByVal sender As Object = Nothing, Optional ByVal e As System.ComponentModel.CancelEventArgs = Nothing)
        Dim sender As Object
        CheckObjectState = False
        Try
            If Not obj Is Nothing Then
                If obj.IsDirty Then
                    Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                    If Results = MsgBoxResult.Yes Then
                        obj.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Function
                        End If

                        If Not bolValidateSuccess Then
                            Exit Function
                        End If
                        If obj.GetType.ToString.ToLower = "muster.businesslogic.ppersona" Then
                            UIUtilsGen.SetOwnerName(Me)
                        End If
                    Else
                        If Results = MsgBoxResult.No Then
                            Dim evt As System.EventArgs
                            If obj.GetType.ToString.ToLower = "muster.businesslogic.ppersona" Then
                                Me.btnOwnerNameCancel_Click(sender, evt)
                            End If
                        End If
                        If Results = MsgBoxResult.Cancel Then
                            Dim e As System.ComponentModel.CancelEventArgs
                            If Not obj.GetType.ToString.ToLower = "muster.businesslogic.ppersona" Then
                                e.Cancel = True
                            Else
                                CheckObjectState = True
                            End If

                        End If

                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub ClearCheckBox(ByVal objControl As Control)
        Dim tmpChk As CheckBox
        Dim currentControl As Control
        Try
            Dim myEnumerator As System.Collections.IEnumerator = _
                           objControl.Controls.GetEnumerator()
            While myEnumerator.MoveNext()
                currentControl = myEnumerator.Current
                If currentControl.GetType.ToString.ToLower = "system.Windows.Forms.checkbox".ToLower Then
                    tmpChk = CType(currentControl, System.Windows.Forms.CheckBox)
                    tmpChk.Checked = False
                    tmpChk.Tag = String.Empty
                    tmpChk.Text = String.Empty
                    'Else
                    '    If currentControl.Controls.Count > 0 Then
                    '        ClearCheckBox(currentControl)
                    '    End If
                End If

            End While
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Closure Event Operations"
#Region "Lookup Methods"
    Private Sub PopulateClosureEventLookups()
        Try
            PopulateClosureType()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateClosureType()
        Try
            cmbClosureType.DisplayMember = "PROPERTY_NAME"
            cmbClosureType.ValueMember = "PROPERTY_ID"
            Me.cmbClosureType.DataSource = pClosure.PopulateClosureType
            Me.cmbClosureType.SelectedIndex = -1
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateClosureFillMaterial()
        Try
            If cmbFillMaterial.Enabled Then
                cmbFillMaterial.DisplayMember = "PROPERTY_NAME"
                cmbFillMaterial.ValueMember = "PROPERTY_ID"
                cmbFillMaterial.DataSource = pClosure.PopulateFillMaterial
                cmbFillMaterial.SelectedIndex = -1
            Else
                cmbFillMaterial.DataSource = Nothing
            End If
            cmbClosureReportFillMaterial.DisplayMember = "PROPERTY_NAME"
            cmbClosureReportFillMaterial.ValueMember = "PROPERTY_ID"
            Me.cmbClosureReportFillMaterial.DataSource = pClosure.PopulateFillMaterial
            Me.cmbClosureReportFillMaterial.SelectedIndex = -1
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "UI Support Routines"
    Private Sub InitializeLetterCount()
        slLetterCount = New SortedList
        slLetterCount.Add(pClosure.LetterType.CIPApproval, 0)
        slLetterCount.Add(pClosure.LetterType.CIPDisApproval, 0)
        slLetterCount.Add(pClosure.LetterType.CIPNFA, 0)
        slLetterCount.Add(pClosure.LetterType.InfoNeeded, 0)
        slLetterCount.Add(pClosure.LetterType.RFGApproval, 0)
        slLetterCount.Add(pClosure.LetterType.RFGNFA, 0)
        slLetterCount.Add(pClosure.LetterType.SampleResultMemo, 0)
    End Sub
    Private Sub ClearClosureForm()
        Try
            bolLoading = True
            UIUtilsGen.ClearFields(tbPageClosures)
            ClearCheckBox(Me.pnlChecklistDetails)
            lblStatusValue.Text = String.Empty
            bolLoading = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub setClosureSaveCancel(ByVal bolValue As Boolean)
        Try
            Me.btnSaveClosure.Enabled = bolValue
            Me.btnClosureCancel.Enabled = bolValue
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetupAddClosureEventForm()
        Try


            RemoveHandler pClosure.evtClosureEventInfoChanged, AddressOf pClosure_evtClosureEventInfoChanged

            nCurrentEventID = -1
            If Not tbCntrlClosure.TabPages.Contains(tbPageClosures) Then
                tbCntrlClosure.TabPages.Add(tbPageClosures)
            End If
            tbCntrlClosure.SelectedTab = tbPageClosures

            ' make all panels visible except for checklist
            lblNoticeOfInterestDisplay.Text = "-"
            PnlNoticeOfInterestDetails.Visible = True
            lblTanksPipesDisplay.Text = "-"
            PnlTanksPipesDetails.Visible = True
            lblAnalysisDisplay.Text = "-"
            pnlAnalysisDetails.Visible = True
            lblChecklistDisplay.Text = "+"
            pnlChecklistDetails.Visible = False
            lblClosureReportDisplay.Text = "-"
            pnlClosureReportDetails.Visible = True
            'lblContactsDisplay.Text = "-"
            'pnlContactDetails.Visible = True
            pnlContacts.Visible = False
            Me.pnlContactDetails.Visible = False
            lblContactsDisplay.Text = "+"

            ClearClosureForm()
            bolLoading = True
            PopulateClosureEventLookups()
            ClearComboboxValues()
            Me.SetUpCheckList()
            SetUpdtPickDateClosed()
            lblClosureIDValue.Visible = False
            EnableDisableClosureControls(False)
            Me.chkShowAllTanksPipes.Checked = True
            pnlChecklistDetails.Visible = False
            lblChecklistDisplay.Text = "+"
            bolLoading = False
            LoadContacts(ugClosureContacts, 0, 22)
            chkClosureShowContactsForAllModules.Checked = False

            lnkLblPrevClosure.Visible = False
            lnkLblNextClosure.Visible = False
            Me.ListAnyScheduledInspections()

            AddHandler pClosure.evtClosureEventInfoChanged, AddressOf pClosure_evtClosureEventInfoChanged


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetupModifyClosureEventForm(ByVal nClosureId As Integer)
        Try
            nCurrentEventID = nClosureId
            If Not tbCntrlClosure.TabPages.Contains(tbPageClosures) Then
                tbCntrlClosure.TabPages.Add(tbPageClosures)
            End If
            Me.getClosureEventsGrid()
            LoadClosureData(nClosureId)
            bolLoading = True
            Me.SetUpCheckList()
            SetUpdtPickDateClosed()
            bolLoading = False
            EnableDisableClosureControls(True)
            Me.chkShowAllTanksPipes.Checked = True

            pnlContacts.Visible = True
            'Me.pnlContactDetails.Visible = True
            'lblContactsDisplay.Text = "-"
            If pClosure.NOIProcessed Then
                Me.btnProcessNOI.Text = "regenerate NOI Letter"
                Me.btnProcessNOI.Enabled = True
            End If

            If pClosure.ClosureProcessed Then
                btnProcessClosure.Text = "Regenerate Process letter"
                btnProcessClosure.Enabled = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub




    Private Sub ListAnyScheduledInspections()

        Dim ins As New DataAccess.InspectionDB
        Dim ds As DataSet = ins.DBGetDS(String.Format("exec sp_ListfacilitiesToBeInspected {0}", Me.pOwn.Facilities.ID))


        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            TxtInspections.Text = String.Format("{0}     Inspector: {1}", Convert.ToDateTime(row("MinScheduled")).ToShortDateString, row("Inspector"))
        End If

        ds = Nothing
        ins = Nothing
    End Sub




    Private Sub LoadClosureData(ByVal nClosure_Id As Integer)

        RemoveHandler pClosure.evtClosureEventInfoChanged, AddressOf pClosure_evtClosureEventInfoChanged

        'Dim sCheckedList As New SortedList
        'Dim sDtClosedList As New SortedList
        'Dim drRow As DataRow
        Dim i As Integer = 0
        Dim objCompany As New MUSTER.BusinessLogic.pCompany
        strAnalysisType = String.Empty
        alAnalysisType = New ArrayList
        Try
            'If Not sCheckedList Is Nothing Then
            '    sCheckedList.Clear()
            'End If
            'If Not sDtClosedList Is Nothing Then
            '    sDtClosedList.Clear()
            'End If
            nCurrentEventID = nClosure_Id
            pClosure.Retrieve(pOwn.Facility, nClosure_Id)

            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
            InitializeLetterCount()

            If Not IsNothing(pClosure) Then
                bolLoading = True
                'sCheckedList = pClosure.BoolCheckList
                'sDtClosedList = pClosure.DateCheckList
                bolLoading = False

                lblClosureIDValue.Tag = pClosure.ID
                Me.Tag = lblClosureIDValue.Tag
                lblClosureIDValue.Visible = True
                lblClosureIDValue.Text = pClosure.FacilitySequence
                lblClosureIDValue.Left = lblClosureID.Left + lblClosureID.Width
                tbCntrlClosure.SelectedTab = tbPageClosures
                ClearClosureForm()
                bolLoading = True
                Me.PopulateClosureEventLookups()
                bolLoading = False
                'If Not pClosure.ClosureType > 0 Then
                '    pnlChecklistDetails.Visible = False
                '    lblChecklistDisplay.Text = "+"
                'End If
                ' to hide the checklist panel
                pnlChecklistDetails.Visible = False
                lblChecklistDisplay.Text = "+"

                UIUtilsGen.SetComboboxItemByValue(cmbClosureType, pClosure.ClosureType)

                If pClosure.NOIReceived = 1 Then
                    rbtnNOIReceivedYes.Checked = True
                    rbtnNOIReceivedNo.Checked = False
                ElseIf pClosure.NOIReceived = 0 Then
                    rbtnNOIReceivedYes.Checked = False
                    rbtnNOIReceivedNo.Checked = True
                Else
                    rbtnNOIReceivedYes.Checked = False
                    rbtnNOIReceivedNo.Checked = False
                End If
                pClosure.CheckNOIReceived()

                lblStatusValue.Text = pClosure.ClosureStatusDesc
                If pClosure.ClosureStatus = 869 Then
                    Me.btnReopen.Enabled = True
                Else
                    Me.btnReopen.Enabled = False
                End If
                UIUtilsGen.SetDatePickerValue(dtPickReceived, pClosure.NOI_Rcv_Date.ToShortDateString)
                UIUtilsGen.SetDatePickerValue(dtPickScheduledDate, pClosure.ScheduledDate.ToShortDateString)


                UIUtilsGen.SetComboboxItemByValue(cmbFillMaterial, pClosure.FillMaterial)
                cmbVerbalWaiver.Checked = pClosure.VerbalWaiver
                UIUtilsGen.SetDatePickerValue(dtPickDueBy, pClosure.DueDate)
                Me.chkShowAllTanksPipes.Checked = True


                UIUtilsGen.SetDatePickerValue(dtPickClosureReceived, pClosure.CRClosureReceived)
                UIUtilsGen.SetDatePickerValue(dtPickDateClosed, pClosure.CRClosureDate)
                UIUtilsGen.SetDatePickerValue(dtPickDateLastUsed, pClosure.CRDateLastUsed)

                Me.GetClosureTanksAndPipesForFacility(pClosure.FacilityID, pClosure.ID)

                ListAnyScheduledInspections()


                RefreshAnalysisGrid(False)
                'For Each drRow In pClosure.SamplesTable.Rows
                '    dtAnalysis.ImportRow(drRow)
                '    If Not alAnalysisType.Contains(drRow.Item("Analysis Type")) Then
                '        alAnalysisType.Add(drRow.Item("Analysis Type"))
                '        strAnalysisType += drRow.Item("Analysis Type") + ","
                '    End If
                'Next
                'FillAnalysisDataGrid()

                Me.txtClosureReportCompany.Text = objCompany.Retrieve(pClosure.CRCompany).COMPANY_NAME()
                pLicensee.Retrieve(pClosure.CRCertContractor)
                txtLicensee.Text = pLicensee.Licensee_name

                txtNOICompany.Text = objCompany.Retrieve(pClosure.Company).COMPANY_NAME()
                pLicensee.Retrieve(pClosure.CertContractor)
                txtNOILicensee.Text = pLicensee.Licensee_name

                UIUtilsGen.SetComboboxItemByValue(cmbClosureReportFillMaterial, pClosure.FillMaterial)
                lblOwnerLastEditedBy.Text = "Last Edited By : " & IIf(pClosure.ModifiedBy = String.Empty, pClosure.CreatedBy.ToString, pClosure.ModifiedBy.ToString)
                lblOwnerLastEditedOn.Text = "Last Edited On : " & IIf(pClosure.ModifiedOn = CDate("01/01/0001"), pClosure.CreatedOn.ToString, pClosure.ModifiedOn.ToString)
                'Load Contacts
                LoadContacts(ugClosureContacts, pClosure.ID, 22)
                chkClosureShowContactsForAllModules.Checked = False

                If nClosure_Id > 0 Then
                    CommentsMaintenance(, , True)
                Else
                    CommentsMaintenance(, , True, True)
                End If

                If dgClosureFacilityDetails.Rows.Count > 0 Then
                    lnkLblPrevClosure.Visible = True
                    lnkLblNextClosure.Visible = True
                Else
                    lnkLblPrevClosure.Visible = False
                    lnkLblNextClosure.Visible = False
                End If
            End If
            ' Me.cmbClosureType.Focus()



            AddHandler pClosure.evtClosureEventInfoChanged, AddressOf pClosure_evtClosureEventInfoChanged
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub ClearComboboxValues()
        bolLoading = True
        Try
            ClearAnalysisDataGrid()
            'dtAnalysis = New DataTable
            SetupAnalysisDataTable()
            ugTankandPipes.DataSource = Nothing
            cmbClosureType.SelectedIndex = -1
            cmbFillMaterial.SelectedIndex = -1
            cmbClosureReportFillMaterial.SelectedIndex = -1
            bolLoading = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub GetClosureTanksAndPipesForFacility(ByVal facID As Integer, ByVal closureID As Integer)
        Dim dtTanksPipes As DataTable
        'Dim strTanksPipes() As String
        'Dim strEntity() As String
        'Dim alTanksPipes As New ArrayList
        'Dim nTankEntityID As Integer
        'Dim nPipeEntityID As Integer
        'Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        'Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        'Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugTankandPipes.DataSource = pClosure.ClosureTankPipeDataSet(facID, closureID)
            ugTankandPipes.Rows.ExpandAll(True)

            ugTankandPipes.DisplayLayout.Bands(0).Columns("PREVIOUS_SUBSTANCE").Hidden = True
            ugTankandPipes.DisplayLayout.Bands(0).Columns("CLOSURE_ID").Hidden = True
            ugTankandPipes.DisplayLayout.Bands(0).Columns("FACILITY_ID").Hidden = True
            ugTankandPipes.DisplayLayout.Bands(0).Columns("TANK_ID").Hidden = True
            ugTankandPipes.DisplayLayout.Bands(0).Columns("CLOSURE_TANK_PIPE_ID").Hidden = True
            ugTankandPipes.DisplayLayout.Bands(0).Columns("POSITION").Hidden = True

            For g As Integer = 1 To 2

                ugTankandPipes.DisplayLayout.Bands(g).Columns("CLOSURE_ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(g).Columns("FACILITY_ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(g).Columns("TANK_ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(g).Columns("PIPE_ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(g).Columns("PREVIOUS_SUBSTANCE").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(g).Columns("COMPARTMENT_NUMBER").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(g).Columns("CLOSURE_TANK_PIPE_ID").Hidden = True
                ugTankandPipes.DisplayLayout.Bands(g).Columns("POSITION").Hidden = True

            Next g
            ugTankandPipes.DisplayLayout.Bands(2).Columns("Parent_Pipe_ID").Hidden = True


            If ugTankandPipes.Rows.Count > 0 Then
                ugTankandPipes.DisplayLayout.Bands(0).Columns("TANK SITE ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(0).Columns("STATUS").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(0).Columns("CLOSURE TYPE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(0).Columns("CAPACITY").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugTankandPipes.DisplayLayout.Bands(0).Columns("MATERIAL").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                For g As Integer = 1 To 2
                    ugTankandPipes.DisplayLayout.Bands(g).Columns("PIPE SITE ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ugTankandPipes.DisplayLayout.Bands(g).Columns("STATUS").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ugTankandPipes.DisplayLayout.Bands(g).Columns("CLOSURE TYPE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ugTankandPipes.DisplayLayout.Bands(g).Columns(" ").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    'ugTankandPipes.DisplayLayout.Bands(g).Columns(" ").CellAppearance.BackColor = SystemColors.Control
                    ugTankandPipes.DisplayLayout.Bands(g).Columns("MATERIAL").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                Next g


                Dim strCloTankPipeList As String = pClosure.GetClosureTankPipeList(closureID)
                If strCloTankPipeList.Trim <> String.Empty Then
                    pClosure.TankPipeID = strCloTankPipeList.Trim.Split("~")(0)
                    pClosure.TankPipeEntity = strCloTankPipeList.Trim.Split("~")(1)
                Else
                    pClosure.TankPipeID = String.Empty
                    pClosure.TankPipeEntity = String.Empty
                End If
            Else
                pClosure.TankPipeID = String.Empty
                pClosure.TankPipeEntity = String.Empty
            End If

            ' disable and change color to gray if tank / pipe already associated with closure event
            ' of type rfg
            For Each ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugTankandPipes.Rows

                Dim isTOSI As Boolean = False

                If ugrow.Cells("STATUS").Value.ToString.ToUpper = "TEMPORARILY OUT OF SERVICE INDEFINITELY" Then
                    isTOSI = True
                End If

                If Not (ugrow.Cells("Last Used").Value Is DBNull.Value) AndAlso ugrow.Cells("last Used").Value > dtMaxDate AndAlso isTOSI Then
                    dtMaxDate = ugrow.Cells("Last Used").Value
                End If

                If ugrow.Cells("CLOSURE TYPE").Text.ToString.IndexOf("Removed from Ground") >= 0 Then
                    If ugrow.Cells("CLOSURE_ID").Value Is DBNull.Value Then
                        'ugrow.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        ugrow.Appearance.BackColor = Color.Gray
                    Else
                        If ugrow.Cells("CLOSURE_ID").Value <> pClosure.ID Then
                            'ugrow.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                            ugrow.Appearance.BackColor = Color.Gray
                        End If
                    End If
                End If
                For Each ugrowChild As Infragistics.Win.UltraWinGrid.UltraGridRow In ugrow.ChildBands(0).Rows

                    If Not (ugrow.Cells("Last Used").Value Is DBNull.Value) AndAlso ugrow.Cells("last Used").Value > dtMaxDate AndAlso isTOSI Then
                        dtMaxDate = ugrow.Cells("Last Used").Value
                    End If

                    If ugrowChild.HasChild Then

                        ugrowChild.Appearance.BackColor = Color.LightGoldenrodYellow

                        For Each ugrowchild2 As Infragistics.Win.UltraWinGrid.UltraGridRow In ugrowChild.ChildBands(0).Rows

                            ugrowchild2.Appearance.BackColor = Color.LightSkyBlue
                        Next

                    End If

                    If ugrowChild.Cells("CLOSURE TYPE").Text.ToString.IndexOf("Removed from Ground") >= 0 Then
                        If ugrowChild.Cells("CLOSURE_ID").Value Is DBNull.Value Then
                            'ugrowChild.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                            ugrowChild.Appearance.BackColor = Color.Gray
                        Else
                            If ugrowChild.Cells("CLOSURE_ID").Value <> pClosure.ID Then
                                'ugrowChild.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                                ugrowChild.Appearance.BackColor = Color.Gray
                            End If
                        End If
                    End If
                Next
            Next

            'If pClosure.TankPipeID <> String.Empty Then
            '    strTanksPipes = pClosure.TankPipeID.Split("|")
            '    strEntity = pClosure.TankPipeEntity.Split("|")
            '    For i As Integer = 0 To strTanksPipes.GetUpperBound(0)
            '        alTanksPipes.Add(strTanksPipes(i) + "|" + strEntity(i))
            '    Next
            '    If alTanksPipes Is Nothing Then
            '        Dim str As String
            '        str = ""
            '    ElseIf alTanksPipes.Count < 1 Then
            '        Dim str As String
            '        str = ""
            '    End If
            '    bolLoading = True
            '    nTankEntityID = uiutilsgen.EntityTypes.Tank
            '    nPipeEntityID = uiutilsgen.EntityTypes.Pipe

            '    For Each ugrow In ugTankandPipes.Rows
            '        If alTanksPipes.Contains(ugrow.Cells("TANK_ID").Value.ToString + "|" + nTankEntityID.ToString) Then
            '            ugrow.Cells("Included").Value = True
            '        End If

            '        ' child rows
            '        For Each ugrowChild As Infragistics.Win.UltraWinGrid.UltraGridRow In ugrow.ChildBands(0).Rows
            '            If alTanksPipes.Contains(ugrowChild.Cells("PIPE_ID").Value.ToString + "|" + nPipeEntityID.ToString) Then
            '                ugrowChild.Cells("Included").Value = True
            '            End If
            '        Next

            '    Next

            '    bolLoading = False
            '    Me.chkShowAllTanksPipes.Checked = True
            '    Me.chkShowAllTanksPipes.Checked = False
            'End If
            If pClosure.ID <= 0 Then
                chkShowAllTanksPipes.Checked = False
                chkShowAllTanksPipes.Checked = True
            Else
                If Not bolCloEventTankPipeGridUpdateInProcess Then
                    chkShowAllTanksPipes.Checked = False
                    chkShowAllTanksPipes.Checked = True
                End If
            End If

            'ugTankandPipes.ActiveCell = ugTankandPipes.Rows(0).Cells("Included")
            'ugTankandPipes.Focus()
            'ugTankandPipes.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ToggleCellSel, False, False)
            'ugTankandPipes.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)

            ValidatePreviousSubstance()

            If pClosure.ID <= 0 Then
                Me.ugTankandPipes.Enabled = False
            Else
                Me.ugTankandPipes.Enabled = True
            End If

            If (Me.dtPickDateLastUsed.Value < "1/1/1930" OrElse Not Me.dtPickDateLastUsed.Checked) AndAlso Me.dtMaxDate > "1/1/1930" Then
                Me.dtPickDateLastUsed.Checked = True
                Me.dtPickDateLastUsed.Value = Me.dtMaxDate
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub EnableDisableClosureControls(ByVal bolState As Boolean)
        Try
            Me.btnDeleteClosure.Enabled = bolState
            Me.btnClosureComments.Enabled = bolState
            Me.btnProcessClosure.Enabled = bolState
            Me.btnProcessNOI.Enabled = bolState
            Me.btnProcessNOIEnvelopes.Enabled = bolState
            If bolState And pClosure.NOIReceived = 0 Then
                Me.btnProcessNOI.Enabled = Not bolState
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Sub SetUpdtPickDateClosed()
        Try
            AddHandler dtPickClosures1.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures2.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures3.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures4.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures5.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures6.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures7.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures8.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures9.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures10.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures11.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures12.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures13.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures14.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures15.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures16.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures17.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures18.ValueChanged, AddressOf dtPick_ValueChanged
            AddHandler dtPickClosures19.ValueChanged, AddressOf dtPick_ValueChanged
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub dtPick_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If bolLoading Then Exit Sub
        Dim dtPick As DateTimePicker
        Try
            dtPick = CType(sender, DateTimePicker)
            Dim dtPickTag As Integer = dtPick.Tag
            UIUtilsGen.ToggleDateFormat(dtPick)
            dtPick.Tag = dtPickTag
            If dtPick.Checked = False Then
                pClosure.UpdateDateCheckList(dtPick.Tag, CDate("01/01/0001"))
            Else
                pClosure.UpdateDateCheckList(dtPick.Tag, UIUtilsGen.GetDatePickerValue(dtPick).Date)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Private Sub setValue(ByVal objControl As Control, ByVal nTag As Integer, ByVal dtValue As DateTime, Optional ByVal chkChecked As Boolean = False)
    '    Try
    '        Dim currentControl As Control
    '        Dim tmpCmb As System.Windows.Forms.CheckBox
    '        Dim tmpDtPick As System.Windows.Forms.DateTimePicker
    '        Dim myEnumerator As System.Collections.IEnumerator = _
    '       objControl.Controls.GetEnumerator()

    '        While myEnumerator.MoveNext()
    '            currentControl = myEnumerator.Current
    '            If currentControl.GetType.ToString.ToLower = "system.Windows.Forms.checkbox".ToLower Then
    '                tmpCmb = CType(currentControl, System.Windows.Forms.CheckBox)
    '                If Not dtValue = Nothing Then
    '                    If tmpCmb.Tag = nTag Then
    '                        pClosure.UpdatechkBoxDate(tmpCmb.Tag, tmpCmb.Checked, dtValue.ToShortDateString)
    '                    End If

    '                End If
    '            ElseIf currentControl.GetType.ToString.ToLower = "system.Windows.Forms.DateTimePicker".ToLower Then
    '                tmpDtPick = CType(currentControl, System.Windows.Forms.DateTimePicker)
    '                If dtValue = Nothing Then
    '                    If tmpDtPick.Tag = nTag Then
    '                        If tmpDtPick.Checked = False Then
    '                            pClosure.UpdatechkBoxDate(nTag, chkChecked)
    '                        Else
    '                            pClosure.UpdatechkBoxDate(nTag, chkChecked, tmpDtPick.Value.ToShortDateString)
    '                        End If
    '                    End If

    '                End If
    '            End If
    '        End While
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    Private Sub SetUpCheckList()
        Try
            AddHandler chkClosures1.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures2.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures3.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures4.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures5.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures6.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures7.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures8.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures9.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures10.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures11.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures12.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures13.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures14.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures15.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures16.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures17.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures18.CheckedChanged, AddressOf chkCheckList_CheckedChanged
            AddHandler chkClosures19.CheckedChanged, AddressOf chkCheckList_CheckedChanged
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub chkCheckList_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If bolLoading Then Exit Sub
        Dim chkBox As CheckBox
        Try
            chkBox = CType(sender, CheckBox)
            pClosure.UpdateBoolCheckList(chkBox.Tag, chkBox.Checked)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub populateCheckList(Optional ByVal nClosureType As Integer = 0, Optional ByVal sCheckedList As SortedList = Nothing, Optional ByVal sDtClosedList As SortedList = Nothing)
    '    Dim ctrl As Control
    '    Dim i As Integer = 0
    '    'Dim j As Integer = 0
    '    Dim chkCol As New Collection
    '    Dim dtPickCol As New Collection
    '    Dim sListChk As New SortedList
    '    Dim sListDt As New SortedList
    '    Dim dtValue As DateTime

    '    Try
    '        If pClosure.ClosureType = 443 Or pClosure.ClosureType <= 0 Then
    '            pnlChecklistDetails.Visible = False
    '            'pClosure.BoolCheckList = Nothing
    '            'pClosure.DateCheckList = Nothing
    '            pClosure.ClearChecklist()
    '            'pClosure.BoolCheckList.Clear()
    '            'pClosure.DateCheckList.Clear()
    '            Exit Sub
    '        End If
    '        For Each ctrl In pnlChecklistDetails.Controls
    '            If ctrl.GetType.ToString.ToLower = "System.Windows.Forms.CheckBox".ToLower Then
    '                chkCol.Add(ctrl, ctrl.Name)
    '            End If
    '            If ctrl.GetType.ToString.ToLower = "System.Windows.Forms.DateTimePicker".ToLower Then
    '                dtPickCol.Add(ctrl, ctrl.Name)
    '            End If
    '        Next
    '        Dim checkListItems As SortedList
    '        If nClosureType > 0 Then
    '            checkListItems = pClosure.PopulateCheckListItems(nClosureType)
    '        Else
    '            checkListItems = sCheckedList
    '        End If
    '        For i = 0 To checkListItems.Count - 1 Step 1

    '            For Each chkbox As CheckBox In chkCol
    '                If chkbox.Name = "chkClosures" + (i + 1).ToString Then
    '                    'If nClosureType > 0 Then
    '                    chkbox.Tag = checkListItems.GetKey(i)
    '                    chkbox.Text = checkListItems.GetByIndex(i)
    '                    If Not pClosure.BoolCheckList Is Nothing Then
    '                        If pClosure.BoolCheckList.Count > 0 And i <= pClosure.BoolCheckList.Count Then
    '                            chkbox.Tag = checkListItems.GetKey(i)
    '                            chkbox.Checked = pClosure.BoolCheckList.GetByIndex(i)
    '                        End If
    '                    End If
    '                    sListChk.Add(chkbox.Tag, chkbox.Checked)

    '                    'Else
    '                    'chkbox.Tag = checkListItems.GetKey(i)
    '                    'chkbox.Checked = checkListItems.GetByIndex(i)
    '                    'sListChk.Add(chkbox.Tag, chkbox.Checked)

    '                    'End If
    '                    Exit For
    '                End If
    '            Next
    '            For Each dtPick As DateTimePicker In dtPickCol
    '                If dtPick.Name = "dtPickClosures" + (i + 1).ToString Then
    '                    'If nClosureType > 0 Then
    '                    dtPick.Tag = checkListItems.GetKey(i)
    '                    'sListDt.Add(dtPick.Tag, CDate("01/01/0001"))
    '                    'Else
    '                    'For j = i To sDtClosedList.Count - 1 Step 1
    '                    If Not pClosure.DateCheckList Is Nothing Then
    '                        If pClosure.DateCheckList.Count > 0 And i <= pClosure.DateCheckList.Count Then
    '                            dtPick.Tag = pClosure.DateCheckList.GetKey(i)
    '                            If pClosure.DateCheckList.GetByIndex(i) Is System.DBNull.Value Then
    '                                dtValue = CDate("01/01/0001")
    '                            Else
    '                                dtValue = CType(pClosure.DateCheckList.GetByIndex(i), DateTime)
    '                            End If
    '                            'dtPick.Tag = sDtClosedList.GetKey(i)
    '                            'If sDtClosedList.GetByIndex(i) Is System.DBNull.Value Then
    '                            '    dtValue = CDate("01/01/0001")
    '                            'Else
    '                            '    dtValue = CType(sDtClosedList.GetByIndex(i), DateTime)
    '                            'End If
    '                            UIUtilsGen.SetDatePickerValue(dtPick, dtValue)
    '                        End If
    '                    End If
    '                    sListDt.Add(dtPick.Tag, dtValue)
    '                    'Exit For
    '                    'Next
    '                    'End If
    '                    Exit For
    '                End If
    '            Next

    '        Next
    '        'If nClosureType > 0 Then
    '        pClosure.BoolCheckList = sListChk
    '        pClosure.DateCheckList = sListDt
    '        'End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    Private Sub populateCheckList(Optional ByVal nClosureType As Integer = 0, Optional ByVal htCheckedList As Hashtable = Nothing, Optional ByVal htDtClosedList As Hashtable = Nothing)
        Dim ctrl As Control
        Dim slchkCol As New SortedList
        Dim sldtPickCol As New SortedList
        Dim dtValue As DateTime

        Try
            If pClosure.ClosureType <> 444 And pClosure.ClosureType <> 445 Then
                pnlChecklistDetails.Visible = False
                pClosure.ClearChecklist()
                Exit Sub
            End If

            Dim checkListItems As Hashtable
            If nClosureType > 0 Then
                checkListItems = pClosure.PopulateCheckListItems(nClosureType)
            Else
                checkListItems = htCheckedList
            End If

            ' reset the tags
            chkClosures1.Tag = "1"
            chkClosures2.Tag = "2"
            chkClosures3.Tag = "3"
            chkClosures4.Tag = "4"
            chkClosures5.Tag = "5"
            chkClosures6.Tag = "6"
            chkClosures7.Tag = "7"
            chkClosures8.Tag = "8"
            chkClosures9.Tag = "9"
            chkClosures10.Tag = "10"
            chkClosures11.Tag = "11"
            chkClosures12.Tag = "12"
            chkClosures13.Tag = "13"
            chkClosures14.Tag = "14"
            chkClosures15.Tag = "15"
            chkClosures16.Tag = "16"
            chkClosures17.Tag = "17"
            chkClosures18.Tag = "18"
            chkClosures19.Tag = "19"

            dtPickClosures1.Tag = "1"
            dtPickClosures2.Tag = "2"
            dtPickClosures3.Tag = "3"
            dtPickClosures4.Tag = "4"
            dtPickClosures5.Tag = "5"
            dtPickClosures6.Tag = "6"
            dtPickClosures7.Tag = "7"
            dtPickClosures8.Tag = "8"
            dtPickClosures9.Tag = "9"
            dtPickClosures10.Tag = "10"
            dtPickClosures11.Tag = "11"
            dtPickClosures12.Tag = "12"
            dtPickClosures13.Tag = "13"
            dtPickClosures14.Tag = "14"
            dtPickClosures15.Tag = "15"
            dtPickClosures16.Tag = "16"
            dtPickClosures17.Tag = "17"
            dtPickClosures18.Tag = "18"
            dtPickClosures19.Tag = "19"

            ' rename the chkbox and dtpick to the value in keys
            Dim sl As New SortedList
            For Each intKey As Integer In checkListItems.Keys
                ' item ID
                sl.Add(intKey, intKey)
            Next

            For Each ctrl In pnlChecklistDetails.Controls
                If ctrl.GetType.ToString.ToLower = "System.Windows.Forms.CheckBox".ToLower Then
                    If ctrl.Name.StartsWith("chkClosures") Then
                        slchkCol.Add(CType(ctrl.Tag, Integer), ctrl)
                    End If
                ElseIf ctrl.GetType.ToString.ToLower = "System.Windows.Forms.DateTimePicker".ToLower Then
                    If ctrl.Name.StartsWith("dtPickClosures") Then
                        sldtPickCol.Add(CType(ctrl.Tag, Integer), ctrl)
                    End If
                End If
            Next

            For i As Integer = 0 To slchkCol.Count - 1
                Dim chkBox As CheckBox
                chkBox = slchkCol.GetByIndex(i)
                chkBox.Tag = sl.GetByIndex(i).ToString ' id
                chkBox.Text = checkListItems.Item(sl.GetByIndex(i))
                If Not pClosure.HashTableBoolCheckList Is Nothing Then
                    chkBox.Checked = pClosure.HashTableBoolCheckList.Item(sl.GetByIndex(i))
                End If
                Dim dtPick As DateTimePicker
                dtPick = sldtPickCol.GetByIndex(i)
                dtPick.Tag = sl.GetByIndex(i).ToString
                If Not pClosure.HashTableDateCheckList Is Nothing Then
                    If pClosure.HashTableDateCheckList.Item(sl.GetByIndex(i)) Is System.DBNull.Value Then
                        dtValue = CDate("01/01/0001")
                    Else
                        dtValue = CType(pClosure.HashTableDateCheckList.Item(sl.GetByIndex(i)), DateTime)
                    End If
                    UIUtilsGen.SetDatePickerValue(dtPick, dtValue)
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub CancelAllOverdue()
        Try
            If pClosure.ID <= 0 Then
                pClosure.CreatedBy = MusterContainer.AppUser.ID
            Else
                pClosure.ModifiedBy = MusterContainer.AppUser.ID
            End If

            pClosure.CancelAllOverdue(pOwn.Facility, CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub GridClosureSelected(ByRef ThisGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        If Me.Cursor Is Cursors.WaitCursor Then
            Exit Sub
        End If
        Dim SelRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            Me.Cursor = Cursors.WaitCursor
            SelRow = ThisGrid.ActiveRow
            Me.SetupModifyClosureEventForm(Integer.Parse(SelRow.Cells("NOI ID").Text))

        Catch ex As Exception
            Throw ex
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub UpdateTankPipe()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugchildrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim pTank As MUSTER.BusinessLogic.pTank
        Dim pPipe As MUSTER.BusinessLogic.pPipe
        Dim bolTankModified, bolCreateTODO As Boolean
        Dim bolCreatedTODO As Boolean = False
        Try
            If pOwn.Facilities.Facility.ID <> pClosure.FacilityID Then
                pOwn.Facilities.Retrieve(pOwn.OwnerInfo, pClosure.FacilityID, , "FACILITY")
            End If
            pTank = pOwn.Facilities.FacilityTanks
            pPipe = pOwn.Facilities.FacilityTanks.Compartments.Pipes
            bolCreateTODO = False
            For Each ugrow In ugTankandPipes.Rows
                bolTankModified = False
                If ugrow.Cells("Included").Value = True Then
                    bolTankModified = True
                    pTank.Retrieve(pOwn.Facilities.Facility, pClosure.FacilityID, Integer.Parse(ugrow.Cells("TANK_ID").Value))
                    pTank.TankStatus = UIUtilsGen.TankPipeStatus.POU
                    pTank.ClosureType = pClosure.ClosureType
                    pTank.DateClosureReceived = pClosure.CRClosureReceived
                    pTank.DateLastUsed = IIf(pTank.DateLastUsed <= "1/1/1950" AndAlso pTank.TankStatus <> 429, pClosure.CRDateLastUsed, pTank.DateLastUsed)
                    pTank.DateClosed = pClosure.CRClosureDate
                    If pClosure.ClosureType = 445 Then
                        pTank.InertMaterial = pClosure.FillMaterial
                    End If
                    pTank.ModifiedBy = MusterContainer.AppUser.ID
                    pTank.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, , False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    'If Not pTank.Save(MusterContainer.AppUser.ID, "Closure", True) Then
                    '    Return False
                    'End If
                End If
                For Each ugchildrow In ugrow.ChildBands(0).Rows
                    If ugchildrow.Cells("Included").Value = True Then
                        If ugchildrow.ParentRow.Cells("Included").Value = False Then bolCreateTODO = True
                        If pTank.TankId <> ugchildrow.Cells("TANK_ID").Value Then
                            pTank.Retrieve(pOwn.Facilities.Facility, pClosure.FacilityID, Integer.Parse(ugchildrow.Cells("TANK_ID").Value))
                        End If
                        pPipe.Retrieve(pTank.TankInfo, ugchildrow.Cells("TANK_ID").Value.ToString + "|" + ugchildrow.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + ugchildrow.Cells("PIPE_ID").Value.ToString)
                        pPipe.PipeStatusDesc = 426
                        pPipe.ClosureType = pClosure.ClosureType
                        pPipe.DateClosureRecd = pClosure.CRClosureReceived
                        pPipe.DateLastUsed = pClosure.CRDateLastUsed
                        pPipe.DateClosed = pClosure.CRClosureDate
                        If pClosure.ClosureType = 445 Then
                            pPipe.InertMaterial = pClosure.FillMaterial
                        End If
                        If pPipe.PipeID <= 0 Then
                            pPipe.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            pPipe.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        pPipe.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                        'If Not pPipe.Save(True) Then
                        '    Return False
                        'End If
                    ElseIf bolTankModified Then
                        ' changing pipe status of unselected pipes to tosi if attached to selected tank
                        ' change pipe status if status is not pou and not associated with other closure event
                        If ugchildrow.Cells("STATUS").Value.ToString <> "Permanently Out of Use" Then
                            If ugchildrow.Cells("CLOSURE TYPE").Text = String.Empty Or ugchildrow.Cells("CLOSURE TYPE").Text = "N/A" Then
                                bolTankModified = False
                                If pTank.TankId <> ugchildrow.Cells("TANK_ID").Value Then
                                    pTank.Retrieve(pOwn.Facilities.Facility, pClosure.FacilityID, Integer.Parse(ugchildrow.Cells("TANK_ID").Value))
                                End If
                                pPipe.Retrieve(pTank.TankInfo, ugchildrow.Cells("TANK_ID").Value.ToString + "|" + ugchildrow.Cells("COMPARTMENT_NUMBER").Value.ToString + "|" + ugchildrow.Cells("PIPE_ID").Value.ToString)
                                pPipe.PipeStatusDesc = 429 ' TOSI
                                pPipe.DateLastUsed = pClosure.CRDateLastUsed
                                If pPipe.PipeID <= 0 Then
                                    pPipe.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    pPipe.ModifiedBy = MusterContainer.AppUser.ID
                                End If
                                pPipe.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If
                            Else
                            End If
                        End If
                    End If
                Next
            Next

            ' # 2824 If cal entry exists for the facility, do not add cal entry
            Dim colCal As MUSTER.Info.CalendarCollection = MusterContainer.pCalendar.RetrieveByOtherID(0, 0, "Facility : " + pClosure.FacilityID.ToString + " Upcoming Pipe Replacement", "DESCRIPTION")
            If Not colCal Is Nothing Then
                If colCal.Count > 0 Then
                    bolCreateTODO = False
                End If
            End If

            If bolCreatedTODO Then
                ' create to do cal entry for registration group if pipe is closed but its associated tank is not closed
                MusterContainer.pCalendar.Add(New MUSTER.Info.CalendarInfo(0, _
                                                Today.Date, _
                                                Today.Date, _
                                                0, _
                                                "Facility : " + pClosure.FacilityID.ToString + " Upcoming Pipe Replacement", _
                                                "", _
                                                "SYSTEM", _
                                                "Registration", _
                                                False, _
                                                True, _
                                                False, _
                                                False, _
                                                MusterContainer.AppUser.ID, _
                                                Now, _
                                                String.Empty, _
                                                CDate("01/01/0001"), _
                                                UIUtilsGen.EntityTypes.ClosureEvent, _
                                                pClosure.ID))
                MusterContainer.pCalendar.Save()

                Dim mc As MusterContainer = Me.MdiParent
                mc.RefreshCalendarInfo()
                mc.LoadDueToMeCalendar()
                mc.LoadToDoCalendar()
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Function GetPrevNextClosure(ByVal closureID As Integer, ByVal getNext As Boolean) As Integer
        Try
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim retVal As Integer = nCurrentEventID
            Dim sl As New SortedList
            Dim slRow As New SortedList

            For Each ugRow In dgClosureFacilityDetails.Rows
                sl.Add(ugRow.Cells("NOI ID").Value, ugRow.Cells("NOI ID").Value)
                slRow.Add(ugRow.Cells("NOI ID").Value, ugRow)
            Next

            retVal = GetPrevNext(sl, getNext, closureID)
            Return retVal
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
#End Region
#Region "UI Control Events"
    Private Sub lnkLblPrevClosure_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevClosure.LinkClicked
        SetupModifyClosureEventForm(GetPrevNextClosure(nCurrentEventID, False))
    End Sub
    Private Sub lnkLblNextClosure_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextClosure.LinkClicked
        SetupModifyClosureEventForm(GetPrevNextClosure(nCurrentEventID, True))
    End Sub
    Private Sub btnProcessClosure_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProcessClosure.Click
        Try
            bolFromNOI = False
            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
            InitializeLetterCount()

            If pClosure.ID <= 0 Then
                pClosure.CreatedBy = MusterContainer.AppUser.ID
            Else
                pClosure.ModifiedBy = MusterContainer.AppUser.ID
            End If

            If pClosure.ClosureProcessed Then

            End If

            Dim curstatus As Boolean = pClosure.ClosureProcessed

            If pClosure.ProcessClosure(MusterContainer.pCalendar, MusterContainer.AppUser.ID, CType(UIUtilsGen.ModuleID.Closure, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal, , Me.BolAll_AAL_HasBackFillOnly) AndAlso Not curstatus Then

                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                ' change tank / pipe status to pou
                If pClosure.ClosureStatus = 868 Then 'closed
                    pClosure.Save(UIUtilsGen.ModuleID.Closure, MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    Dim ds As DataSet
                    Dim strFees As String = String.Empty
                    ds = pOwn.RunSQLQuery("SELECT dbo.udfGetOwnerPastDueFees(" + pOwn.ID.ToString + ",0," + pClosure.FacilityID.ToString + ")")
                    If ds.Tables(0).Rows(0)(0) > 0 Then
                        strFees = ds.Tables(0).Rows(0)(0)
                        strFees = strFees.Split(".")(0)
                    End If

                    If strFees.Length = 0 Then
                        strFees = "0"
                    End If

                    ' create to do cal entry for Fee group if tank is closed.
                    If CInt(strFees) > 0 Then
                        MusterContainer.pCalendar.Add(New MUSTER.Info.CalendarInfo(0, _
                                                        Today.Date, _
                                                        Today.Date, _
                                                        0, _
                                                        "Facility : " + pClosure.FacilityID.ToString + " Owes Fees - Tank/Pipe Closed", _
                                                        "", _
                                                        MusterContainer.AppUser.RetrieveFEEHead.ID, _
                                                        "Fee Users", _
                                                        False, _
                                                        True, _
                                                        False, _
                                                        False, _
                                                        MusterContainer.AppUser.ID, _
                                                        Now, _
                                                        String.Empty, _
                                                        CDate("01/01/0001"), _
                                                        UIUtilsGen.EntityTypes.ClosureEvent, _
                                                        pClosure.ID))
                        MusterContainer.pCalendar.Save()
                        Dim mc As MusterContainer = Me.MdiParent
                        mc.RefreshCalendarInfo()
                        mc.LoadDueToMeCalendar()
                        mc.LoadToDoCalendar()
                    End If

                    UpdateTankPipe()
                    Me.GetClosureTanksAndPipesForFacility(pClosure.FacilityID, pClosure.ID)
                End If

                If pClosure.ClosureProcessed Then
                    Me.btnProcessClosure.Enabled = False
                End If
                Me.lblStatusValue.Text = pClosure.ClosureStatusDesc
                MsgBox("Process Closure Success")
            Else
                If pClosure.ClosureStatus = 866 Then ' pending
                    Me.lblStatusValue.Text = pClosure.ClosureStatusDesc
                End If
            End If
            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
            InitializeLetterCount()
            'If bolValidateSuccess Then
            '    Me.lblStatusValue.Text = pClosure.ClosureStatusDesc
            '    MsgBox("Process Closure Success")
            'Else
            '    bolValidateSuccess = True
            '    bolDisplayErrmessage = True
            '    nErrMessage = 0
            '    Exit Sub
            'End If
            If Not bolValidateSuccess Then
                bolValidateSuccess = True
                bolDisplayErrmessage = True
                nErrMessage = 0
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnProcessNOI_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProcessNOI.Click
        Try
            'If pClosure.ProcessNOI(MusterContainer.AppUser.ID) Then
            '    pClosure.NOIProcessed = True
            '    If pClosure.Save(True) Then
            '        Me.btnProcessNOI.Enabled = False
            bolFromNOI = True
            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
            InitializeLetterCount()

            If pClosure.ID <= 0 Then
                pClosure.CreatedBy = MusterContainer.AppUser.ID
            Else
                pClosure.ModifiedBy = MusterContainer.AppUser.ID
            End If

            pClosure.ProcessNOI(MusterContainer.pCalendar, MusterContainer.AppUser.ID, CType(UIUtilsGen.ModuleID.Closure, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
            InitializeLetterCount()
            If pClosure.NOIProcessed Then
                Me.btnProcessNOI.Enabled = False
            End If
            If bolValidateSuccess Then
                Me.lblStatusValue.Text = pClosure.ClosureStatusDesc
                MsgBox("Process NOI Success")
            Else
                bolValidateSuccess = True
                bolDisplayErrmessage = True
                nErrMessage = 0
                Exit Sub
            End If
            bolFromNOI = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAddClosure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddClosure.Click
        Try
            SetupAddClosureEventForm()
            pClosure.Retrieve(pOwn.Facility, 0)
            bolLoading = True

            UIUtilsGen.SetDatePickerValue(dtPickReceived, pClosure.NOI_Rcv_Date.ToShortDateString)
            UIUtilsGen.SetDatePickerValue(dtPickScheduledDate, pClosure.ScheduledDate.ToShortDateString)
            bolLoading = False
            GetClosureTanksAndPipesForFacility(pClosure.FacilityID, pClosure.ID)
            RefreshAnalysisGrid(True)
            'FillAnalysisDataGrid()
            Me.cmbClosureType.Focus()
            ' pClosure.FacilityID = pOwn.Facilities.ID
            Me.Text = "Closure - Manage Closure (" & lblClosureIDValue.Text & ")"

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkShowAllTanksPipes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAllTanksPipes.CheckedChanged
        Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim bolActiveRow As Boolean
        Try
            If Not chkShowAllTanksPipes.Checked Then
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
                        If ugrow.Cells("INCLUDED").Value = 0 And bolActiveRow = False Then
                            ugrow.Hidden = True
                        End If
                    End If
                Next
            Else
                For Each ugrow In ugTankandPipes.Rows
                    ChildBand = ugrow.ChildBands(0)
                    For Each Childrow In ChildBand.Rows
                        Childrow.Hidden = False
                    Next
                    ugrow.Hidden = False
                Next
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnEvtTankCollapse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEvtTankCollapse.Click
        Try
            If btnEvtTankCollapse.Text = "Collapse All" Then
                ugTankandPipes.Rows.CollapseAll(True)
                btnEvtTankCollapse.Text = "Expand All"
            Else
                ugTankandPipes.Rows.ExpandAll(True)
                btnEvtTankCollapse.Text = "Collapse All"
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbClosureType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbClosureType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pClosure.ClosureType = UIUtilsGen.GetComboBoxValueInt(cmbClosureType)
            bolLoading = True
            populateCheckList(pClosure.ClosureType)
            bolLoading = False

            'If pClosure.ClosureType = 445 Then
            '    Me.cmbFillMaterial.Enabled = True
            '    Me.cmbClosureReportFillMaterial.Enabled = True
            '    PopulateClosureFillMaterial()
            '    populateCheckList(pClosure.ClosureType)
            'Else
            '    Me.cmbFillMaterial.Enabled = False
            '    Me.cmbClosureReportFillMaterial.Enabled = False
            'End If

            ' commented so the checklist panel is hidden until user wants it open
            'If pClosure.ClosureType = 444 Or pClosure.ClosureType = 445 Then
            '    pnlChecklistDetails.Visible = True
            '    lblChecklistDisplay.Text = "-"
            'Else
            '    pnlChecklistDetails.Visible = False
            '    lblChecklistDisplay.Text = "+"
            'End If

            Me.btnClosureCancel.Focus()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbClosureReportFillMaterial_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbClosureReportFillMaterial.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pClosure.FillMaterial = UIUtilsGen.GetComboBoxValueInt(cmbClosureReportFillMaterial)
            If cmbFillMaterial.Enabled Then
                cmbFillMaterial.SelectedValue = pClosure.FillMaterial
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub cmbFillMaterial_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFillMaterial.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            pClosure.FillMaterial = UIUtilsGen.GetComboBoxValueInt(cmbFillMaterial)
            cmbClosureReportFillMaterial.SelectedValue = pClosure.FillMaterial
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnReopen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReopen.Click
        Try
            If pClosure.ID <= 0 Then
                pClosure.CreatedBy = MusterContainer.AppUser.ID
            Else
                pClosure.ModifiedBy = MusterContainer.AppUser.ID
            End If

            pClosure.ReopenCancelled(pClosure.ID, CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            lblStatusValue.Text = pClosure.ClosureStatusDesc
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSaveClosure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveClosure.Click
        Dim nClosureID As Integer

        Try

            nClosureID = pClosure.ID

            If pClosure.ID <= 0 Then
                pClosure.CreatedBy = MusterContainer.AppUser.ID
            Else
                pClosure.ModifiedBy = MusterContainer.AppUser.ID
            End If


            Me.bolValidateSuccess = pClosure.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If Me.bolValidateSuccess Then
                If nClosureID <= 0 Then
                    MsgBox("Closure was added successfully!")
                    Me.EnableDisableClosureControls(True)
                Else
                    MsgBox("Closure was updated successfully")
                End If
                If Not pClosure.ClosureType = 443 Then
                    pClosure.ManageChecklist(MusterContainer.pCalendar, MusterContainer.AppUser.ID, pClosure.DueDate, pClosure.DueDate, "Missing Information for Closure NOI # " + pClosure.FacilitySequence.ToString + " on Facility ID " + pClosure.FacilityID.ToString, False, True, String.Empty)
                    Dim MyFrm As MusterContainer
                    MyFrm = Me.MdiParent
                    MyFrm.DisplayOnDateSelectedOrChangedOrViewEntries()
                End If
                SetupModifyClosureEventForm(pClosure.ID)

                ugTankandPipes.Rows.ExpandAll(True)
            Else
                bolValidateSuccess = True
                bolDisplayErrmessage = True
                nErrMessage = 0
                Exit Sub
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClosureCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClosureCancel.Click
        Try

            If Not pClosure Is Nothing Then
                pClosure.Reset()

                'Me.ClearClosureForm()
                'Me.ClearComboboxValues()

                If pClosure.ID > 0 Then
                    SetupModifyClosureEventForm(pClosure.ID)
                    'Me.LoadClosureData(pClosure.ID)
                Else
                    Me.SetupAddClosureEventForm()
                    bolLoading = True
                    UIUtilsGen.SetDatePickerValue(dtPickReceived, pClosure.NOI_Rcv_Date.ToShortDateString)
                    UIUtilsGen.SetDatePickerValue(dtPickScheduledDate, pClosure.ScheduledDate.ToShortDateString)
                    bolLoading = False
                    GetClosureTanksAndPipesForFacility(pClosure.FacilityID, pClosure.ID)
                    RefreshAnalysisGrid(True)
                    Me.cmbClosureType.Focus()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDeleteClosure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteClosure.Click
        Dim result As MsgBoxResult
        Try
            If pClosure.ID <> 0 Then

                Dim deletedClosureID As Integer = pClosure.ID

                result = MessageBox.Show("Are you Sure you want to Delete this Record?", "Closure", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then
                    Exit Sub
                End If
                If pClosure.IsDirty Then
                    result = MessageBox.Show("There are unsaved changes. Do you want to save the changes before Deletion? ", "Closure", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If result = DialogResult.No Then
                        pClosure.Reset()
                    End If
                End If

                If pClosure.ID <= 0 Then
                    pClosure.CreatedBy = MusterContainer.AppUser.ID
                Else
                    pClosure.ModifiedBy = MusterContainer.AppUser.ID
                End If

                If pClosure.DeleteClosure(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal) Then

                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    Dim mc As MusterContainer = Me.MdiParent

                    ' delete all cal entries associated with this event
                    If deletedClosureID > 0 Then
                        Dim colCal As MUSTER.Info.CalendarCollection
                        Dim cal As New MUSTER.BusinessLogic.pCalendar
                        colCal = cal.RetrieveByOtherID(UIUtilsGen.EntityTypes.ClosureEvent, deletedClosureID, , "")
                        For Each calInfo As MUSTER.Info.CalendarInfo In colCal.Values
                            cal.Add(calInfo)
                            cal.Deleted = True
                            cal.Save()
                        Next
                        If colCal.Count > 0 Then
                            If Not mc Is Nothing Then
                                mc.RefreshCalendarInfo()
                                mc.LoadDueToMeCalendar()
                                mc.LoadToDoCalendar()
                            End If
                        End If
                    End If

                    ' Delete associated Flags
                    Dim oFlag As New MUSTER.BusinessLogic.pFlag
                    oFlag.RetrieveFlags(pClosure.FacilityID, UIUtilsGen.EntityTypes.Facility, , , , , "SYSTEM", "Missing Information for Closure NOI # " + pClosure.FacilitySequence.ToString)
                    For Each flagInfo As MUSTER.Info.FlagInfo In oFlag.FlagsCol.Values
                        flagInfo.Deleted = True
                    Next
                    If oFlag.FlagsCol.Count > 0 Then oFlag.Flush()

                    If Not mc Is Nothing Then
                        mc.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text)
                    End If

                    Me.SetupAddClosureEventForm()
                    Me.getClosureEventsGrid()
                    tbCntrlClosure.SelectedTab = tbPageFacilityDetail
                Else
                    nErrMessage = 0
                End If
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickReceived_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickReceived.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickReceived)
            UIUtilsGen.FillDateobjectValues(pClosure.NOI_Rcv_Date, dtPickReceived.Text)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickClosureReceived_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickClosureReceived.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickClosureReceived)
            pClosure.CRClosureReceived = UIUtilsGen.GetDatePickerValue(dtPickClosureReceived)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickDateClosed_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickDateClosed.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickDateClosed)
            pClosure.CRClosureDate = UIUtilsGen.GetDatePickerValue(dtPickDateClosed)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickDateLastUsed_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickDateLastUsed.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickDateLastUsed)
            pClosure.CRDateLastUsed = UIUtilsGen.GetDatePickerValue(dtPickDateLastUsed)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickDueBy_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickDueBy.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickDueBy)
            UIUtilsGen.FillDateobjectValues(pClosure.DueDate, dtPickDueBy.Text)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickScheduledDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickScheduledDate.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickScheduledDate)
            UIUtilsGen.FillDateobjectValues(pClosure.ScheduledDate, dtPickScheduledDate.Text)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub rbtnNOIReceivedNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnNOIReceivedNo.Click
        If bolLoading Then Exit Sub
        Try
            If rbtnNOIReceivedNo.Checked Then
                pClosure.NOIReceived = 0
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub rbtnNOIReceivedYes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnNOIReceivedYes.Click
        If bolLoading Then Exit Sub
        Try
            If rbtnNOIReceivedYes.Checked Then
                pClosure.NOIReceived = 1
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub txtCompany_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If bolLoading Then Exit Sub
    '    Try
    '        If txtCompany.Text <> String.Empty Then
    '            pClosure.Company = txtCompany.Text
    '            txtClosureReportCompany.Text = txtCompany.Text
    '        Else
    '            pClosure.Company = 0
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub txtContact_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If bolLoading Then Exit Sub
    '    Try
    '        If txtContact.Text <> String.Empty Then
    '            pClosure.Contact = txtContact.Text
    '        Else
    '            pClosure.Contact = 0
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub txtClosureReportCompany_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If bolLoading Then Exit Sub
        Try
            If txtClosureReportCompany.Text <> String.Empty Then
                pClosure.CRCompany = txtClosureReportCompany.Text
            Else
                pClosure.CRCompany = 0
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub txtClosureReportCertifiedContractor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If bolLoading Then Exit Sub
    '    Try
    '        If txtClosureReportCertifiedContractor.Text <> String.Empty Then
    '            pClosure.CRCertContractor = txtClosureReportCertifiedContractor.Text
    '        Else
    '            pClosure.CRCertContractor = 0
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub txtCertifiedContractor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If bolLoading Then Exit Sub
    '    Try
    '        If txtCertifiedContractor.Text <> String.Empty Then
    '            pClosure.CertContractor = txtCertifiedContractor.Text
    '        Else
    '            pClosure.CertContractor = 0
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    'Private Sub cmbCertifiedContractor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If bolLoading Then Exit Sub
    '    Try
    '        If cmbCertifiedContractor.SelectedIndex = -1 Then Exit Sub

    '        pClosure.CertContractor = cmbCertifiedContractor.SelectedValue
    '        Dim dsCompany As DataSet
    '        dsCompany = pClosure.GetCompanyName(cmbCertifiedContractor.SelectedValue)
    '        If Not dsCompany Is Nothing Then
    '            txtCompany.Text = dsCompany.Tables(0).Rows(0).Item("Company_Name")
    '            pClosure.Company = dsCompany.Tables(0).Rows(0).Item("Company_ID")

    '            If txtClosureReportCompany.Text = String.Empty Then
    '                txtClosureReportCompany.Text = dsCompany.Tables(0).Rows(0).Item("Company_Name")
    '                pClosure.CRCompany = dsCompany.Tables(0).Rows(0).Item("Company_ID")
    '            End If
    '        Else
    '            txtCompany.Text = String.Empty
    '            pClosure.Company = 0
    '        End If

    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    Private Sub btnClosureComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClosureComments.Click
        Try
            CommentsMaintenance(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbVerbalWaiver_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbVerbalWaiver.Click
        If bolLoading Then Exit Sub
        Try
            pClosure.VerbalWaiver = cmbVerbalWaiver.Checked
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Ultra Grids"
#Region "Lookup Methods"
    Private Sub PopulateAnalysisLevel(ByRef drow As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Dim dtAnalysisLevel As New DataTable
        Dim vListAnalysisLevel As New Infragistics.Win.ValueList
        Dim type As String = String.Empty
        Dim types() As String

        If Not drow.Cells("Analysis Type") Is Nothing AndAlso Not drow.Cells("Analysis Type").Value Is Nothing Then
            type = Convert.ToString(drow.Cells("Analysis Type").Value)

            If type = "1656" Then
                type = "875&876"
            End If

            types = type.Split("&"c)

        End If

        If Not types Is Nothing AndAlso types.Length > 0 Then

            For Each val As String In types

                dtAnalysisLevel = pClosure.PopulateAnalysisLevel(val.Trim)

                If Not (dtAnalysisLevel Is Nothing) Then

                    For Each dr As DataRow In dtAnalysisLevel.Rows

                        If Not vListAnalysisLevel.FindString(dr.Item("PROPERTY_NAME")) > -1 Then
                            vListAnalysisLevel.ValueListItems.Add(dr.Item("PROPERTY_ID"), dr.Item("PROPERTY_NAME").ToString)
                        End If
                    Next
                End If

                drow.Cells("Analysis Level").ValueList = vListAnalysisLevel

            Next val

        End If

        If vListAnalysisLevel.FindByDataValue(drow.Cells("Analysis Level").Value) Is Nothing Then
            drow.Cells("Analysis Level").Value = DBNull.Value
        End If

    End Sub
    Private Sub PopulateSampleLocation(ByRef drow As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Dim vListSampleLocation As New Infragistics.Win.ValueList
        Dim dtSampleLocation As New DataTable
        If drow.Cells("Analysis Level").Value Is DBNull.Value Then
            dtSampleLocation = pClosure.PopulateSampleLocation(pClosure.ClosureType, 0)
        Else
            dtSampleLocation = pClosure.PopulateSampleLocation(pClosure.ClosureType, drow.Cells("Analysis Level").Value)
        End If
        If Not (dtSampleLocation Is Nothing) Then
            For Each dr As DataRow In dtSampleLocation.Rows
                vListSampleLocation.ValueListItems.Add(dr.Item("PROPERTY_ID"), dr.Item("PROPERTY_NAME").ToString)
            Next
        End If
        drow.Cells("Sample Location").ValueList = vListSampleLocation
        If vListSampleLocation.FindByDataValue(drow.Cells("Sample Location").Value) Is Nothing Then
            drow.Cells("Sample Location").Value = DBNull.Value
        End If
    End Sub

    'Private Function populateAnalysisLevel(ByVal nAnalysisType As Integer) As Infragistics.Win.ValueList
    '    Dim drRow As DataRow
    '    Dim dtAnalysisLevel As DataTable
    '    Dim vListAnalysisLevel As New Infragistics.Win.ValueList
    '    Try
    '        dtAnalysisLevel = pClosure.PopulateAnalysisLevel(nAnalysisType)
    '        For Each drRow In dtAnalysisLevel.Rows
    '            vListAnalysisLevel.ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME").ToString)
    '        Next
    '        Return vListAnalysisLevel
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '    End Try

    'End Function

    'Private Sub PopulateSampleMedia()
    '    Try
    '        Dim drRow As DataRow
    '        Dim dtSampleMedia As DataTable = pClosure.PopulateSampleMedia
    '        dGridAnalysis.DisplayLayout.ValueLists("SampleMediaValue").ValueListItems.Clear()
    '        dGridAnalysis.DisplayLayout.ValueLists("SampleMediaValue").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left
    '        For Each drRow In dtSampleMedia.Rows
    '            dGridAnalysis.DisplayLayout.ValueLists("SampleMediaValue").ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME").ToString)
    '        Next
    '        dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Media").ValueList = dGridAnalysis.DisplayLayout.ValueLists("SampleMediaValue")
    '        dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Media").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
    '        dGridAnalysis.DisplayLayout.ValueLists("SampleMediaValue").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '    End Try
    'End Sub
    'Private Sub PopulateAnalysisType()
    '    Dim drRow As DataRow
    '    Try
    '        Dim dtAnalysisType As DataTable = pClosure.PopulateAnalysisType
    '        dGridAnalysis.DisplayLayout.ValueLists("AnalysisTypeValue").ValueListItems.Clear()
    '        dGridAnalysis.DisplayLayout.ValueLists("AnalysisTypeValue").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left
    '        For Each drRow In dtAnalysisType.Rows
    '            dGridAnalysis.DisplayLayout.ValueLists("AnalysisTypeValue").ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME").ToString)
    '        Next
    '        dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Type").ValueList = dGridAnalysis.DisplayLayout.ValueLists("AnalysisTypeValue")
    '        dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Type").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
    '        dGridAnalysis.DisplayLayout.ValueLists("AnalysisTypeValue").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '    End Try
    'End Sub
    'Private Function PopulateSampleLocation(ByVal nClosureType As Integer, ByVal nAnalysisLevel As Integer) As Infragistics.Win.ValueList
    '    Dim drRow As DataRow
    '    Dim vListSampleLocation As New Infragistics.Win.ValueList
    '    Dim dtSampleLocation As New DataTable
    '    Try
    '        dtSampleLocation = pClosure.PopulateSampleLocation(nClosureType, nAnalysisLevel)
    '        If Not dtSampleLocation Is Nothing Then
    '            For Each drRow In dtSampleLocation.Rows
    '                vListSampleLocation.ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME").ToString)
    '            Next
    '        End If
    '        Return vListSampleLocation
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function
    'Private Sub PopulateAnalysisUnits()
    '    Dim drRow As DataRow
    '    Try
    '        Dim dtAnalysisUnits As DataTable = pClosure.PopulateAnalysisUnits
    '        dGridAnalysis.DisplayLayout.ValueLists("AnalysisUnitsValue").ValueListItems.Clear()
    '        dGridAnalysis.DisplayLayout.ValueLists("AnalysisUnitsValue").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left
    '        For Each drRow In dtAnalysisUnits.Rows
    '            dGridAnalysis.DisplayLayout.ValueLists("AnalysisUnitsValue").ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME").ToString)
    '        Next
    '        dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Units").ValueList = dGridAnalysis.DisplayLayout.ValueLists("AnalysisUnitsValue")
    '        dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Units").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
    '        dGridAnalysis.DisplayLayout.ValueLists("AnalysisUnitsValue").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '    End Try
    'End Sub
    'Private Sub PopulateCertifiedContractors()
    '    Try
    '        Me.cmbCertifiedContractor.DataSource = pClosure.PopulateCertifiedContractor
    '        Me.cmbCertifiedContractor.DisplayMember = "Licensee_Name"
    '        Me.cmbCertifiedContractor.ValueMember = "Licensee_ID"
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

#End Region
#Region "UI Support Routines"
    Private Sub SaveSampleTableRow(ByRef row As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            Dim sampleID As Integer = row.Cells("SAMPLE_ID").Value
            Dim closureID As Integer = row.Cells("CLOSURE_ID").Value
            Dim sampleNumber As String = row.Cells("Sample #").Value
            Dim analysisType As Integer = IIf(row.Cells("Analysis Type").Value Is DBNull.Value, 0, row.Cells("Analysis Type").Value)
            Dim analysisLevel As Integer = IIf(row.Cells("Analysis Level").Value Is DBNull.Value, 0, row.Cells("Analysis Level").Value)
            Dim sampleMedia As Integer = IIf(row.Cells("Sample Media").Value Is DBNull.Value, 0, row.Cells("Sample Media").Value)
            Dim sampleLocation As Integer = IIf(row.Cells("Sample Location").Value Is DBNull.Value, 0, row.Cells("Sample Location").Value)
            Dim sampleValue As Single = IIf(row.Cells("Sample Value").Value Is DBNull.Value, 0.0, row.Cells("Sample Value").Value)
            Dim sampleUnits As Integer = IIf(row.Cells("Sample Units").Value Is DBNull.Value, 0, row.Cells("Sample Units").Value)
            Dim sampleConstituent As String = IIf(row.Cells("Sample Constituent").Value Is DBNull.Value, String.Empty, row.Cells("Sample Constituent").Value)
            Dim deleted As Boolean = row.Cells("DELETED").Value

            Dim userID As String = MusterContainer.AppUser.ID
            pClosure.PutClosureSample(sampleID, closureID, sampleNumber, analysisType, analysisLevel, sampleMedia, sampleLocation, sampleValue, sampleUnits, sampleConstituent, deleted, userID, CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If Not deleted And sampleID <= 0 Then
                row.Cells("SAMPLE_ID").Value = sampleID
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SaveTankPipeTableRow(ByVal row As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            Dim cloTankPipeID As Integer = 0
            Dim tankPipeID As Integer = 0
            Dim tankPipeEntity As Integer = 0
            If row.Band.Index = 0 Then
                tankPipeID = row.Cells("TANK_ID").Value
                tankPipeEntity = UIUtilsGen.EntityTypes.Tank
                nLastTankPipeId = tankPipeID
                sLastTankPipeType = "TANK"
            Else
                tankPipeID = row.Cells("PIPE_ID").Value
                tankPipeEntity = UIUtilsGen.EntityTypes.Pipe
                nLastTankPipeId = tankPipeID
                sLastTankPipeType = "PIPE"
            End If
            Dim closureID As Integer = IIf(row.Cells("CLOSURE_ID").Value Is DBNull.Value, pClosure.ID, row.Cells("CLOSURE_ID").Value)
            Dim deleted As Boolean = False

            If row.Cells("INCLUDED").Value = True And row.Cells("INCLUDED").Text = False AndAlso TypeOf row.Cells("CLOSURE_TANK_PIPE_ID").Value Is Integer Then
                cloTankPipeID = row.Cells("CLOSURE_TANK_PIPE_ID").Value
                deleted = True
            End If

            Dim analysisType As Integer = -1
            Dim analysisLevel As Integer = -1
            Dim sampleMedia As Integer = -1
            Dim sampleResultsID As Integer = -1

            pClosure.PutClosureTankPipe(cloTankPipeID, tankPipeID, tankPipeEntity, closureID, deleted, CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, analysisType, analysisLevel, sampleMedia, sampleResultsID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub RefreshAnalysisGrid(Optional ByVal bolgetFromDB As Boolean = False)
        ClearAnalysisDataGrid()
        If bolgetFromDB Then pClosure.PopulateSampleTable(pClosure.ID)
        ' clone the datatable
        SetupAnalysisDataTable()
        BolAll_AAL_HasBackFillOnly = PopulateAnalysisGrid()
        FillAnalysisDataGrid()
        SetupAnalysisGrid()
    End Sub
    Private Sub ClearAnalysisDataGrid()
        Try
            dtAnalysis = New DataTable
            dGridAnalysis.DataSource = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function PopulateAnalysisGrid() As Boolean
        Dim drRow As DataRow
        Dim strAnalysisTypeName As String = String.Empty
        Dim oProperty As New MUSTER.BusinessLogic.pProperty
        Dim hasWater As Boolean = False
        Dim AALCnt As Integer = 0
        Dim AALBackFillOnlyCnt As Integer = 0
        Dim hasSoil As Boolean = False
        Try
            alAnalysisType.Clear()
            strAnalysisType = String.Empty
            For Each drRow In pClosure.SamplesTable.Rows
                dtAnalysis.ImportRow(drRow)
                If Not drRow.Item("Analysis Type") Is DBNull.Value Then
                    If Not drRow.Item("Analysis Type") = 0 Then
                        strAnalysisTypeName = oProperty.GetPropertyNameByID(drRow.Item("Analysis Type"))
                        If Not alAnalysisType.Contains(strAnalysisTypeName) Then
                            alAnalysisType.Add(strAnalysisTypeName)
                            strAnalysisType += strAnalysisTypeName + ","

                        End If
                    End If
                End If
            Next

            For Each drRow In pClosure.SamplesTable.Rows
                If drRow.Item("Analysis Level") = 890 Then
                    AALCnt += 1


                    If Not drRow.Item("Sample Media") Is DBNull.Value AndAlso drRow.Item("Sample Location") = 873 Then
                        AALBackFillOnlyCnt += 1
                    End If
                End If
                If Not drRow.Item("Sample Media") Is DBNull.Value AndAlso CStr(drRow.Item("Sample Media")) = "879" Then
                    hasWater = True
                ElseIf Not drRow.Item("Sample Media") Is DBNull.Value AndAlso CStr(drRow.Item("Sample Media")) = "878" Then
                    hasSoil = True
                End If



            Next



            If hasWater And hasSoil Then

                strMedia = " soil/ water"
            ElseIf hasWater Then
                strMedia = " water"
            Else
                strMedia = " soil"
            End If


            If strAnalysisType.Length > 0 Then strAnalysisType = strAnalysisType.Trim.TrimEnd(",")

            Return (AALCnt > 0 AndAlso AALCnt = AALBackFillOnlyCnt)
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub FillAnalysisDataGrid()
        '  Dim i As Integer = 0
        Try
            'dtAnalysis = pClosure.SamplesTable
            dGridAnalysis.DataSource = Nothing
            dtAnalysis.DefaultView.Sort = "Sample #"
            dGridAnalysis.DataSource = dtAnalysis
            SetupAnalysisGrid()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetupAnalysisGrid()
        Try
            dGridAnalysis.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True
            dGridAnalysis.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom
            dGridAnalysis.DisplayLayout.Override.TemplateAddRowAppearance.BackColor = Color.White

            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Value").MaskInput = "nnnn.nnn"
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Value").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Value").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample #").Width = 90
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Type").Width = 100
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Level").Width = 100
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Media").Width = 100
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Location").Width = 110
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Value").Width = 100
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Units").Width = 100
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Constituent").Width = 120

            dGridAnalysis.DisplayLayout.Bands(0).Columns("CLOSURE_ID").Hidden = True
            dGridAnalysis.DisplayLayout.Bands(0).Columns("SAMPLE_ID").Hidden = True
            dGridAnalysis.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
            dGridAnalysis.DisplayLayout.Bands(0).Columns("CREATED_BY").Hidden = True
            dGridAnalysis.DisplayLayout.Bands(0).Columns("DATE_CREATED").Hidden = True
            dGridAnalysis.DisplayLayout.Bands(0).Columns("LAST_EDITED_BY").Hidden = True
            dGridAnalysis.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True

            dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Type").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Level").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Media").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Location").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Units").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

            If dGridAnalysis.DisplayLayout.ValueLists.All.Length = 0 Then
                dGridAnalysis.DisplayLayout.ValueLists.Add("AnalysisType")
                dGridAnalysis.DisplayLayout.ValueLists.Add("AnalysisLevel")
                dGridAnalysis.DisplayLayout.ValueLists.Add("SampleMedia")
                dGridAnalysis.DisplayLayout.ValueLists.Add("SampleLocation")
                dGridAnalysis.DisplayLayout.ValueLists.Add("SampleUnits")
            End If

            'dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Type").ValueList
            ''dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Level").ValueList
            'dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Media").ValueList
            ''dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Location").ValueList
            'dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Units").ValueList

            ' populate the whole column as the table is the same for each row

            ' Material Of Construction
            'If dGridAnalysis.DisplayLayout.Bands(0).Columns("MATERIALS").ValueList Is Nothing Then
            '    vListMaterial = New Infragistics.Win.ValueList
            '    For Each row As DataRow In pTank.PopulateTankMaterialOfConstruction.Rows
            '        vListMaterial.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
            '    Next
            '    dGridAnalysis.DisplayLayout.Bands(0).Columns("MATERIALS").ValueList = vListMaterial
            'End If

            ' Sample Media
            If dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Media").ValueList Is Nothing Then
                Dim vListSampleMedia As New Infragistics.Win.ValueList
                For Each dr As DataRow In pClosure.PopulateSampleMedia.Rows
                    vListSampleMedia.ValueListItems.Add(dr.Item("PROPERTY_ID"), dr.Item("PROPERTY_NAME").ToString)
                Next
                dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Media").ValueList = vListSampleMedia
            End If

            ' Analysis Type
            If dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Type").ValueList Is Nothing Then
                Dim vListAnalysisType As New Infragistics.Win.ValueList
                For Each dr As DataRow In pClosure.PopulateAnalysisType.Rows
                    vListAnalysisType.ValueListItems.Add(dr.Item("PROPERTY_ID"), dr.Item("PROPERTY_NAME").ToString)
                Next
                dGridAnalysis.DisplayLayout.Bands(0).Columns("Analysis Type").ValueList = vListAnalysisType
            End If

            ' Sample Units
            If dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Units").ValueList Is Nothing Then
                Dim vListSampleUnits As New Infragistics.Win.ValueList
                For Each dr As DataRow In pClosure.PopulateAnalysisUnits.Rows
                    vListSampleUnits.ValueListItems.Add(dr.Item("PROPERTY_ID"), dr.Item("PROPERTY_NAME").ToString)
                Next
                dGridAnalysis.DisplayLayout.Bands(0).Columns("Sample Units").ValueList = vListSampleUnits
            End If

            'MakeEnterActLikeTab(dGridAnalysis)

            For Each drow As Infragistics.Win.UltraWinGrid.UltraGridRow In dGridAnalysis.Rows
                If Not drow.Cells("Analysis Type") Is System.DBNull.Value Then
                    If drow.Cells("Analysis Type").Value <> "0" And drow.Cells("Analysis Type").Value <> String.Empty Then
                        PopulateAnalysisLevel(drow)
                        'Dim dtAnalysisLevel As New DataTable
                        'Dim vListAnalysisLevel As New Infragistics.Win.ValueList
                        'dtAnalysisLevel = pClosure.PopulateAnalysisLevel(drow.Cells("Analysis Type").Value)
                        'If Not (dtAnalysisLevel Is Nothing) Then
                        '    For Each dr As DataRow In dtAnalysisLevel.Rows
                        '        vListAnalysisLevel.ValueListItems.Add(dr.Item("PROPERTY_ID"), dr.Item("PROPERTY_NAME").ToString)
                        '    Next
                        'End If
                        'drow.Cells("Analysis Level").ValueList = vListAnalysisLevel
                        'If vListAnalysisLevel.FindByDataValue(drow.Cells("Analysis Level").Value) Is Nothing Then
                        '    drow.Cells("Analysis Level").Value = DBNull.Value
                        'End If
                    End If
                End If
                If Not drow.Cells("Analysis level").Value Is System.DBNull.Value Then
                    If drow.Cells("Analysis level").Value <> "0" And drow.Cells("Analysis Type").Value <> String.Empty Then
                        PopulateSampleLocation(drow)
                        'Dim vListSampleLocation As New Infragistics.Win.ValueList
                        'Dim dtSampleLocation As New DataTable
                        'dtSampleLocation = pClosure.PopulateSampleLocation(pClosure.ClosureType, drow.Cells("Analysis Level").Value)
                        'If Not (dtSampleLocation Is Nothing) Then
                        '    For Each dr As DataRow In dtSampleLocation.Rows
                        '        vListSampleLocation.ValueListItems.Add(dr.Item("PROPERTY_ID"), dr.Item("PROPERTY_NAME").ToString)
                        '    Next
                        'End If
                        'drow.Cells("Sample Location").ValueList = vListSampleLocation
                        'If vListSampleLocation.FindByDataValue(drow.Cells("Sample Location").Value) Is Nothing Then
                        '    drow.Cells("Sample Location").Value = DBNull.Value
                        'End If
                    End If
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ValidatePreviousSubstance()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugchildrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            For Each ugrow In ugTankandPipes.Rows
                If ugrow.Cells("PREVIOUS_SUBSTANCE").Text <> String.Empty Then
                    ugrow.Cells("SUBSTANCE").Appearance.BackColor = Color.Yellow
                Else
                    'ugrow.Cells("SUBSTANCE").Appearance.BackColor = Color.White
                    ugrow.Cells("SUBSTANCE").Appearance.BackColor = ugrow.Appearance.BackColor
                End If
                For Each ugchildrow In ugrow.ChildBands(0).Rows
                    If ugchildrow.Cells("PREVIOUS_SUBSTANCE").Text <> String.Empty Then
                        ugchildrow.Cells("SUBSTANCE").Appearance.BackColor = Color.Yellow
                    Else
                        'ugchildrow.Cells("SUBSTANCE").Appearance.BackColor = Color.White
                        ugchildrow.Cells("SUBSTANCE").Appearance.BackColor = ugchildrow.Appearance.BackColor
                    End If
                Next
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetupAnalysisDataTable()
        Try
            dtAnalysis = pClosure.SamplesTable.Clone
            'dtAnalysis.Columns.Add("SAMPLE_ID", GetType(Integer))
            'dtAnalysis.Columns.Add("CLOSURE_ID", GetType(Integer))
            'dtAnalysis.Columns.Add("Sample #", GetType(Integer))
            'dtAnalysis.Columns.Add("Analysis Type")
            'dtAnalysis.Columns.Add("Analysis Level")
            'dtAnalysis.Columns.Add("Sample Media")
            'dtAnalysis.Columns.Add("Sample Location")
            'dtAnalysis.Columns.Add("Sample Value", GetType(Single))
            'dtAnalysis.Columns.Add("Sample Units")
            'dtAnalysis.Columns.Add("Sample Constituent", GetType(Integer))
            'dtAnalysis.Columns.Add("DELETED", GetType(Boolean))
            'dtAnalysis.Columns("DELETED").DefaultValue = False
        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
    Private Sub ConfigureUgPreviousSubstance(ByVal dtPreSub As DataTable, ByVal bolTank As Boolean)
        Try

            If dtPreSub.Rows.Count = 0 Then
                udPreviousSubstance.Visible = False
                Exit Sub
            End If
            Me.udPreviousSubstance.DataSource = dtPreSub
            Me.udPreviousSubstance.Enabled = False

            Me.udPreviousSubstance.ValueMember = "PREVIOUS SUBSTANCE"
            Me.udPreviousSubstance.DisplayMember = "PREVIOUS SUBSTANCE"

            If bolTank = True Then
                ugTankandPipes.DisplayLayout.Bands(0).Columns("SUBSTANCE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                Me.ugTankandPipes.DisplayLayout.Bands(0).Columns("SUBSTANCE").ValueList = Me.udPreviousSubstance
            Else
                ugTankandPipes.DisplayLayout.Bands(1).Columns("SUBSTANCE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                Me.ugTankandPipes.DisplayLayout.Bands(1).Columns("SUBSTANCE").ValueList = Me.udPreviousSubstance
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub displayPreviousSubstance(ByVal strPreSubstance() As String, ByVal bolTank As Boolean)
        Dim dtPreviousSubstance As New DataTable
        Dim drRow As DataRow
        dtPreviousSubstance.Columns.Add("PREVIOUS SUBSTANCE")
        Try
            dtPreviousSubstance.Rows.Clear()
            If Not strPreSubstance Is Nothing Then
                For i As Integer = 0 To UBound(strPreSubstance)
                    If strPreSubstance(i) <> String.Empty Then
                        drRow = dtPreviousSubstance.NewRow
                        drRow("PREVIOUS SUBSTANCE") = strPreSubstance(i)
                        dtPreviousSubstance.Rows.Add(drRow)
                    End If
                Next
            End If
            Me.ConfigureUgPreviousSubstance(dtPreviousSubstance, bolTank)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Private Sub MakeEnterActLikeTab(ByVal Grid As Infragistics.Win.UltraWinGrid.UltraGrid)
    '    Dim ugKam As Infragistics.Win.UltraWinGrid.GridKeyActionMapping
    '    Dim newKam As Infragistics.Win.UltraWinGrid.GridKeyActionMapping
    '    For Each ugKam In Grid.KeyActionMappings
    '        If ugKam.KeyCode = Keys.Tab Then
    '            newKam = New Infragistics.Win.UltraWinGrid.GridKeyActionMapping(Keys.Tab, ugKam.ActionCode, ugKam.StateDisallowed, ugKam.StateRequired, ugKam.SpecialKeysDisallowed, ugKam.SpecialKeysRequired)
    '            Grid.KeyActionMappings.Add(newKam)
    '        End If
    '    Next
    'End Sub
    Private Function ValidData(ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) As Boolean
        If e.NewValue < 0 Or e.NewValue > 9999 Then
            MessageBox.Show("Value must be from 1 to 9999")
            e.Cancel = True
            Return False
        Else
            Return True
        End If
    End Function
    'Private Sub fillSampleAnalysis()
    '    Dim dtSample As New DataTable
    '    Dim drrow As DataRow
    '    Dim dr As DataRow
    '    Try

    '        dtSample = dtAnalysis.Clone
    '        For Each drrow In dtAnalysis.Rows
    '            'dtSample.ImportRow(drrow)
    '            dr = dtSample.NewRow
    '            dr.Item("SAMPLE_ID") = drrow.Item("SAMPLE_ID")
    '            dr.Item("CLOSURE_ID") = drrow.Item("CLOSURE_ID")
    '            dr.Item("Sample #") = drrow.Item("Sample #")
    '            dr.Item("Analysis Type") = drrow.Item("Analysis Type")
    '            dr.Item("Analysis Level") = drrow.Item("Analysis Level")
    '            dr.Item("Sample Media") = drrow.Item("Sample Media")
    '            dr.Item("Sample Location") = drrow.Item("Sample Location")
    '            dr.Item("Sample Value") = drrow.Item("Sample Value")
    '            dr.Item("Sample Units") = drrow.Item("Sample Units")
    '            dr.Item("Sample Constituent") = drrow.Item("Sample Constituent")
    '            dr.Item("DELETED") = drrow.Item("DELETED")
    '            dtSample.Rows.Add(dr)
    '        Next
    '        pClosure.SamplesTable = dtSample
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region
#Region "UI Control Events"
    Private Sub ugTankandPipes_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles ugTankandPipes.BeforeCellActivate
        Dim strPreviousSub() As String
        Try
            If "SUBSTANCE".Equals(e.Cell.Column.Key) Then
                strPreviousSub = Nothing
                If e.Cell.Row.HasParent Then
                    If e.Cell.Row.Cells("PREVIOUS_SUBSTANCE").Text <> String.Empty Then
                        strPreviousSub = e.Cell.Row.Cells("PREVIOUS_SUBSTANCE").Text.Split("|")
                    End If
                    displayPreviousSubstance(strPreviousSub, False)
                    If strPreviousSub Is Nothing Then
                        e.Cancel = True
                    Else
                        e.Cancel = False
                    End If
                Else
                    If e.Cell.Row.Cells("PREVIOUS_SUBSTANCE").Text <> String.Empty Then
                        strPreviousSub = e.Cell.Row.Cells("PREVIOUS_SUBSTANCE").Text.Split("|")
                    End If
                    displayPreviousSubstance(strPreviousSub, True)
                    If strPreviousSub Is Nothing Then
                        e.Cancel = True
                    Else
                        e.Cancel = False
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTankandPipes_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTankandPipes.CellChange
        If bolLoading Then Exit Sub
        If bolCloEventTankPipeGridUpdateInProcess Then Exit Sub
        Dim selRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim cbChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugParentRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        'Dim oEntity As New MUSTER.BusinessLogic.pEntity
        'Dim nTankEntityID As Integer = uiutilsgen.EntityTypes.Tank
        'Dim nPipeEntityID As Integer = uiutilsgen.EntityTypes.Pipe
        'Dim strTanksAndPipes As String = String.Empty
        'Dim strEntityId As String = String.Empty
        'Dim alTankRemoved As New ArrayList
        'Dim alPipeRemoved As New ArrayList
        'Dim alTankAdded As New ArrayList
        'Dim alPipeAdded As New ArrayList
        'Dim alTank As New ArrayList
        'Dim alPipe As New ArrayList
        Try
            ' setting bolCloEventTankPipeGridUpdateInProcess = true so that when you modify the value of a cell, it wont run through the same process again

            If pClosure.ID <= 0 Then
                If MsgBox("Please Save Closure event before attaching Tanks / Pipes" + vbCrLf + _
                        "Do you want to save Closure event?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
                    If pClosure.ID <= 0 Then
                        pClosure.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pClosure.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    If Not pClosure.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal) Then
                        e.Cell.Value = e.Cell.OriginalValue
                        Exit Sub
                    End If

                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                Else
                    e.Cell.Value = e.Cell.OriginalValue
                    Exit Sub
                End If
            End If

            bolCloEventTankPipeGridUpdateInProcess = True

            If "INCLUDED".Equals(e.Cell.Column.Key.ToUpper) Then
                selRow = e.Cell.Row
                If e.Cell.Row.Cells("CLOSURE_ID").Value Is DBNull.Value Then
                    If e.Cell.Row.Cells("CLOSURE TYPE").Text.Trim.ToUpper = "REMOVED FROM GROUND" Then
                        e.Cell.CancelUpdate()
                        Exit Sub
                    End If
                End If
                SaveTankPipeTableRow(e.Cell.Row)
                ' refresh grid and bring focus to the row
                ''                GetClosureTanksAndPipesForFacility(pClosure.FacilityID, pClosure.ID)
                ugTankandPipes.ActiveRow = selRow
                'Set focus on last row changed.
                Dim tmp As Infragistics.Win.UltraWinGrid.UltraGridRow
                If sLastTankPipeType = "TANK" Then
                    With ugTankandPipes
                        'Get the first row in the UltraGrid
                        tmp = .GetRow(Infragistics.Win.UltraWinGrid.ChildRow.First)
                        Do Until tmp.Cells("TANK_ID").Value = nLastTankPipeId
                            If tmp.HasNextSibling Then
                                tmp = tmp.GetSibling(Infragistics.Win.UltraWinGrid.SiblingRow.Next)
                            Else
                                tmp = Nothing
                                Exit Do
                            End If
                        Loop

                        'Check to see if tmp is nothing. If it is, the row was not found.
                        If Not tmp Is Nothing Then
                            'If the row was found set the FirstRow Property of the 
                            'RowScrollRegion to bring it in to view
                            .ActiveRowScrollRegion.FirstRow = tmp
                        End If
                    End With
                ElseIf sLastTankPipeType = "PIPE" Then
                    With ugTankandPipes
                        For Each ugParentRow In .Rows
                            For Each cbChildRow In ugParentRow.ChildBands(0).Rows
                                'Get the first row in the UltraGrid
                                tmp = .GetRow(Infragistics.Win.UltraWinGrid.ChildRow.First)
                                Do Until tmp.Cells("TANK_ID").Value = nLastTankPipeId
                                    If tmp.HasNextSibling Then
                                        tmp = tmp.GetSibling(Infragistics.Win.UltraWinGrid.SiblingRow.Next)
                                    Else
                                        tmp = Nothing
                                        Exit Do
                                    End If
                                Loop

                                'Check to see if tmp is nothing. If it is, the row was not found.
                                If Not tmp Is Nothing Then
                                    'If the row was found set the FirstRow Property of the 
                                    'RowScrollRegion to bring it in to view
                                    .ActiveRowScrollRegion.FirstRow = tmp
                                End If
                            Next
                        Next
                    End With
                End If
            End If

            'If "INCLUDED".Equals(e.Cell.Column.Key.ToUpper) Then
            '    If e.Cell.Value = True And e.Cell.Text = False Then
            '        If e.Cell.Band.Index = 0 Then
            '            If Not alTankRemoved.Contains(e.Cell.Row.Cells("TANK_ID").Value) Then
            '                alTankRemoved.Add(e.Cell.Row.Cells("TANK_ID").Value)
            '            End If
            '        ElseIf e.Cell.Band.Index = 1 Then
            '            If Not alPipeRemoved.Contains(e.Cell.Row.Cells("PIPE_ID").Value) Then
            '                alPipeRemoved.Add(e.Cell.Row.Cells("PIPE_ID").Value)
            '            End If
            '        End If
            '    ElseIf e.Cell.Value = False And e.Cell.Text = True Then
            '        If e.Cell.Band.Index = 0 Then
            '            If Not alTankAdded.Contains(e.Cell.Row.Cells("TANK_ID").Value) Then
            '                alTankAdded.Add(e.Cell.Row.Cells("TANK_ID").Value)
            '            End If
            '        ElseIf e.Cell.Band.Index = 1 Then
            '            If Not alPipeAdded.Contains(e.Cell.Row.Cells("PIPE_ID").Value) Then
            '                alPipeAdded.Add(e.Cell.Row.Cells("PIPE_ID").Value)
            '            End If
            '        End If
            '    End If
            'End If

            'For Each ugrow In ugTankandPipes.Rows
            '    If alTankRemoved.Contains(ugrow.Cells("TANK_ID").Value) And ugrow.Cells("Included").Value = True Then
            '        ugrow.Cells("Included").Value = False
            '    End If
            '    If alTankAdded.Contains(ugrow.Cells("TANK_ID").Value) And ugrow.Cells("Included").Value = False Then
            '        ugrow.Cells("Included").Value = True
            '    End If
            '    If ugrow.Cells("Included").Value = True Then
            '        If Not alTank.Contains(ugrow.Cells("TANK_ID").Value) Then
            '            strEntityId += nTankEntityID.ToString + "|"
            '            strTanksAndPipes += ugrow.Cells("TANK_ID").Value.ToString + "|"
            '            alTank.Add(ugrow.Cells("TANK_ID").Value)
            '        End If
            '    End If
            '    For Each Childrow In ugrow.ChildBands(0).Rows
            '        If alPipeRemoved.Contains(Childrow.Cells("PIPE_ID").Value) And Childrow.Cells("Included").Value = True Then
            '            Childrow.Cells("Included").Value = False
            '        End If
            '        If alPipeAdded.Contains(Childrow.Cells("PIPE_ID").Value) And Childrow.Cells("Included").Value = False Then
            '            Childrow.Cells("Included").Value = True
            '        End If
            '        If Childrow.Cells("Included").Value = True Then
            '            If Not alPipe.Contains(Childrow.Cells("PIPE_ID").Value) Then
            '                strEntityId += nPipeEntityID.ToString + "|"
            '                strTanksAndPipes += Childrow.Cells("PIPE_ID").Value.ToString + "|"
            '                alPipe.Add(Childrow.Cells("PIPE_ID").Value)
            '            End If
            '        End If
            '    Next
            'Next

            'strEntityId = strEntityId.Trim.TrimEnd("|")
            'strTanksAndPipes = strTanksAndPipes.Trim.TrimEnd("|")
            'pClosure.TankPipeID = strEntityId
            'pClosure.TankPipeEntity = strtanksandpipes

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolCloEventTankPipeGridUpdateInProcess = False
        End Try
    End Sub
    'Private Sub ugTankandPipes_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugTankandPipes.AfterCellUpdate
    '    If bolLoading Then Exit Sub
    '    If bolCloEventTankPipeGridUpdateInProcess Then Exit Sub

    '    Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    '    Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    '    'Dim oEntity As New MUSTER.BusinessLogic.pEntity
    '    Dim nTankEntityID As Integer = uiutilsgen.EntityTypes.Tank
    '    Dim nPipeEntityID As Integer = uiutilsgen.EntityTypes.Pipe
    '    Dim strTanksAndPipes As String = String.Empty
    '    Dim strEntityId As String = String.Empty
    '    Dim alTankRemoved As New ArrayList
    '    Dim alPipeRemoved As New ArrayList
    '    Dim alTankAdded As New ArrayList
    '    Dim alPipeAdded As New ArrayList
    '    Dim alTank As New ArrayList
    '    Dim alPipe As New ArrayList
    '    Try
    '        ' setting bolCloEventTankPipeGridUpdateInProcess = true so that when you modify the value of a cell, it wont run through the same process again
    '        bolCloEventTankPipeGridUpdateInProcess = True

    '        If e.Cell.Column.Key.ToUpper = "INCLUDED" Then
    '            If e.Cell.OriginalValue = True And e.Cell.Value = False Then
    '                If e.Cell.Band.Index = 0 Then
    '                    If Not alTankRemoved.Contains(e.Cell.Row.Cells("TANK_ID").Value) Then
    '                        alTankRemoved.Add(e.Cell.Row.Cells("TANK_ID").Value)
    '                    End If
    '                ElseIf e.Cell.Band.Index = 1 Then
    '                    If Not alPipeRemoved.Contains(e.Cell.Row.Cells("PIPE_ID").Value) Then
    '                        alPipeRemoved.Add(e.Cell.Row.Cells("PIPE_ID").Value)
    '                    End If
    '                End If
    '            ElseIf e.Cell.OriginalValue = False And e.Cell.Value = True Then
    '                If e.Cell.Band.Index = 0 Then
    '                    If Not alTankAdded.Contains(e.Cell.Row.Cells("TANK_ID").Value) Then
    '                        alTankAdded.Add(e.Cell.Row.Cells("TANK_ID").Value)
    '                    End If
    '                ElseIf e.Cell.Band.Index = 1 Then
    '                    If Not alPipeAdded.Contains(e.Cell.Row.Cells("PIPE_ID").Value) Then
    '                        alPipeAdded.Add(e.Cell.Row.Cells("PIPE_ID").Value)
    '                    End If
    '                End If
    '            End If
    '        End If

    '        For Each ugrow In ugTankandPipes.Rows
    '            If alTankRemoved.Contains(ugrow.Cells("TANK_ID").Value) And ugrow.Cells("Included").Value = True Then
    '                ugrow.Cells("Included").Value = False
    '            End If
    '            If alTankAdded.Contains(ugrow.Cells("TANK_ID").Value) And ugrow.Cells("Included").Value = False Then
    '                ugrow.Cells("Included").Value = True
    '            End If
    '            If ugrow.Cells("Included").Value = True Then
    '                If Not alTank.Contains(ugrow.Cells("TANK_ID").Value) Then
    '                    strEntityId += nTankEntityID.ToString + "|"
    '                    strTanksAndPipes += ugrow.Cells("TANK_ID").Value.ToString + "|"
    '                    alTank.Add(ugrow.Cells("TANK_ID").Value)
    '                End If
    '            End If
    '            For Each Childrow In ugrow.ChildBands(0).Rows
    '                If alPipeRemoved.Contains(Childrow.Cells("PIPE_ID").Value) And Childrow.Cells("Included").Value = True Then
    '                    Childrow.Cells("Included").Value = False
    '                End If
    '                If alPipeAdded.Contains(Childrow.Cells("PIPE_ID").Value) And Childrow.Cells("Included").Value = False Then
    '                    Childrow.Cells("Included").Value = True
    '                End If
    '                If Childrow.Cells("Included").Value = True Then
    '                    If Not alPipe.Contains(Childrow.Cells("PIPE_ID").Value) Then
    '                        strEntityId += nPipeEntityID.ToString + "|"
    '                        strTanksAndPipes += Childrow.Cells("PIPE_ID").Value.ToString + "|"
    '                        alPipe.Add(Childrow.Cells("PIPE_ID").Value)
    '                    End If
    '                End If
    '            Next
    '        Next

    '        strEntityId = strEntityId.Trim.TrimEnd("|")
    '        strTanksAndPipes = strTanksAndPipes.Trim.TrimEnd("|")

    '        pClosure.TankPipeID = strTanksAndPipes
    '        pClosure.TankPipeEntity = strEntityId

    '        bolCloEventTankPipeGridUpdateInProcess = False
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugTankandPipes_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles ugTankandPipes.InitializeRow
    '    Dim dtPreviousSubstance As New DataTable
    '    Dim nTankPipeID As Integer
    '    Dim nEntityID As Integer

    '    'dtPreviousSubstance = pClosure.GetTankPipePreviousSubstance(pClosure.FacilityID, nTankPipeID, nEntityID)
    '    Select Case e.Row.Cells(1).Column.ToString.ToUpper
    '        Case "TANK SITE ID"
    '            nEntityID = 10
    '            nTankPipeID = Integer.Parse(e.Row.Cells("TANK_ID").Value)
    '            dtPreviousSubstance = pClosure.GetTankPipePreviousSubstance(pClosure.FacilityID, nTankPipeID, nEntityID)
    '            If Not dtPreviousSubstance Is Nothing Then
    '                e.Row.Cells("SUBSTANCE").Appearance.BackColor = Color.Yellow
    '            Else
    '                e.Row.Cells("SUBSTANCE").Appearance.BackColor = Color.White
    '            End If
    '        Case "PIPE SITE ID"
    '            nEntityID = 12
    '            nTankPipeID = Integer.Parse(e.Row.Cells("PIPE_ID").Value)
    '            dtPreviousSubstance = pClosure.GetTankPipePreviousSubstance(pClosure.FacilityID, nTankPipeID, nEntityID)
    '            If Not dtPreviousSubstance Is Nothing Then
    '                e.Row.Cells("SUBSTANCE").Appearance.BackColor = Color.Yellow
    '            Else
    '                e.Row.Cells("SUBSTANCE").Appearance.BackColor = Color.White
    '            End If
    '    End Select
    'End Sub

    Private Sub dGridAnalysis_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dGridAnalysis.KeyPress
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        'Dim ugCell As Infragistics.Win.UltraWinGrid.UltraGridCell

        ugrow = Me.dGridAnalysis.ActiveRow
        'ugCell = dGridAnalysis.ActiveCell
        'If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Or e.KeyChar = Microsoft.VisualBasic.ChrW(127) Then
        '    dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
        'End If
        If Not e.KeyChar = Microsoft.VisualBasic.ChrW(9) Then
            'Select Case ugCell.Column.ToString.ToUpper
            Select Case dGridAnalysis.ActiveCell.Column.ToString.ToUpper
                Case "ANALYSIS TYPE"
                    If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                        dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        dGridAnalysis.ActiveCell.Value = Nothing
                        ugrow.Cells("Analysis Level").Value = Nothing
                        ugrow.Cells("Sample Location").Value = Nothing
                        ugrow.Cells("Sample Media").Value = Nothing
                        ugrow.Cells("Sample Value").Value = 0
                        ugrow.Cells("Sample Units").Value = Nothing
                        ugrow.Cells("Sample Constituent").Value = 0
                    End If
                Case "ANALYSIS LEVEL"
                    If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                        dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        dGridAnalysis.ActiveCell.Value = Nothing
                        ugrow.Cells("Sample Location").Value = Nothing
                    End If

                Case "SAMPLE MEDIA"
                    If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                        dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        dGridAnalysis.ActiveCell.Value = Nothing
                    End If
                Case "SAMPLE LOCATION"
                    If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                        dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        dGridAnalysis.ActiveCell.Value = Nothing
                    End If
                Case "SAMPLE UNITS"
                    If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                        dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                        dGridAnalysis.ActiveCell.Value = Nothing
                    End If
                Case "SAMPLE VALUE"
                    If e.KeyChar.IsNumber(e.KeyChar) = False Then
                        If Not e.KeyChar = Microsoft.VisualBasic.ChrW(8) And Not e.KeyChar = Microsoft.VisualBasic.ChrW(46) Then
                            e.Handled = True
                        End If
                    End If
                    If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Or e.KeyChar = Microsoft.VisualBasic.ChrW(127) Then
                        If dGridAnalysis.ActiveCell.Text.Length > 1 Then
                            dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                        Else
                            dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                            dGridAnalysis.ActiveCell.Value = 0
                        End If
                    End If
                Case "SAMPLE CONSTITUENT"
                    'If e.KeyChar.IsNumber(e.KeyChar) = False Then
                    '    If Not e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                    '        e.Handled = True
                    '    End If
                    'End If
                    If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Or e.KeyChar = Microsoft.VisualBasic.ChrW(127) Then
                        If dGridAnalysis.ActiveCell.Text.Length > 1 Then
                            dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                        Else
                            dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                            dGridAnalysis.ActiveCell.Value = 0
                        End If
                    End If
                Case "SAMPLE #"
                    'If e.KeyChar.IsNumber(e.KeyChar) = False Then
                    '    If Not e.KeyChar = Microsoft.VisualBasic.ChrW(8) Then
                    '        e.Handled = True
                    '    End If
                    'End If
                    If e.KeyChar = Microsoft.VisualBasic.ChrW(8) Or e.KeyChar = Microsoft.VisualBasic.ChrW(127) Then
                        If dGridAnalysis.ActiveCell.Text.Length > 1 Then
                            dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
                        Else
                            dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
                            dGridAnalysis.ActiveCell.Value = 0
                        End If
                    End If
            End Select

        End If
        'If (KeyAscii >= 48 And KeyAscii <= 57) Then
        '    KeyAscii = KeyAscii
        'ElseIf KeyAscii <> 8 Then
        '    KeyAscii = 0

        'End If

    End Sub
    'Private Sub dGridAnalysis_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles dGridAnalysis.BeforeCellUpdate
    '    Select Case e.Cell.Column.ToString.ToUpper
    '        Case "SAMPLE #"
    '            ValidData(e)
    '        Case "SAMPLE VALUE"
    '            'If e.NewValue = 0 Then
    '            '    e.Cancel = True
    '            '    Exit Select
    '            'End If
    '            ValidData(e)
    '        Case "SAMPLE CONSTITUENT"
    '            'If e.NewValue = 0 Then
    '            '    e.Cancel = True
    '            '    Exit Select
    '            'End If
    '            ValidData(e)
    '    End Select

    'End Sub
    'Private Sub dGridAnalysis_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles dGridAnalysis.BeforeCellActivate
    '    Try
    '        If e.Cell.Column.ToString.ToUpper.IndexOf("SAMPLE #") < 0 Then
    '            'If e Then .Cell.Value Is System.DBNull.Value Then
    '            If e.Cell.Row.Cells("Sample #").Value Is System.DBNull.Value Then
    '                MsgBox("Sample # is required")
    '                e.Cancel = True
    '                Exit Sub
    '                'dGridAnalysis.Focus()
    '                'dGridAnalysis.ActiveCell = dGridAnalysis.ActiveRow.Cells("Sample #")
    '                'dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ToggleCellSel, False, False)
    '                'dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
    '                Exit Sub
    '            End If
    '        End If
    '        Select Case e.Cell.Column.ToString.ToUpper
    '            Case UCase("Analysis Level")
    '                'If Not e.Cell.Row.Cells("Analysis Type").Value Is System.DBNull.Value Then
    '                If Not e.Cell.Row.Cells("Analysis Type").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw).ToUpper Is String.Empty Then
    '                    PopulateAnalysisLevel(e.Cell.Row)
    '                    'e.Cell.ValueList = populateAnalysisLevel(Integer.Parse(e.Cell.Row.Cells("Analysis Type").Value))
    '                    e.Cell.Column.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
    '                Else
    '                    e.Cell.Value = Nothing
    '                    e.Cancel = True
    '                End If
    '            Case UCase("Sample Location")
    '                If Not e.Cell.Row.Cells("Analysis Level").Value Is System.DBNull.Value Then
    '                    If e.Cell.Row.Cells("Analysis Level").Value <> String.Empty Then
    '                        PopulateSampleLocation(e.Cell.Row)
    '                        'e.Cell.ValueList = PopulateSampleLocation(pClosure.ClosureType, Integer.Parse(e.Cell.Row.Cells("Analysis Level").Value))
    '                        e.Cell.Column.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
    '                    Else
    '                        e.Cell.Value = Nothing
    '                        e.Cancel = True
    '                    End If
    '                End If

    '        End Select
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub dGridAnalysis_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles dGridAnalysis.AfterCellUpdate
    '    Select Case e.Cell.Column.ToString.ToUpper
    '        Case "ANALYSIS LEVEL"
    '            If e.Cell.Row.Cells("Analysis Level").Text <> String.Empty Then
    '                If e.Cell.Row.Cells("Analysis Level").Value = 884 Or e.Cell.Row.Cells("Analysis Level").Value = 885 Then
    '                    e.Cell.Row.Cells("Sample Location").Value = Nothing
    '                End If
    '                'Else
    '                '    e.Cell.Row.Cells("Analysis Level").Value = Nothing
    '                '    e.Cell.Row.Cells("Sample Location").Value = Nothing
    '            End If
    '            'Case "ANALYSIS TYPE"
    '            '    If e.Cell.Row.Cells("Analysis Type").Text = String.Empty Then
    '            '        e.Cell.Row.Cells("Analysis Level").Value = Nothing
    '            '        e.Cell.Row.Cells("Sample Location").Value = Nothing
    '            '        e.Cell.Row.Cells("Sample Media").Value = Nothing
    '            '        e.Cell.Row.Cells("Sample Value").Value = String.Empty
    '            '        e.Cell.Row.Cells("Sample Units").Value = Nothing
    '            '        e.Cell.Row.Cells("Sample Constituent").Value = String.Empty

    '            '    End If
    '    End Select
    '    fillSampleAnalysis()
    'End Sub
    'Private Sub dGridAnalysis_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dGridAnalysis.AfterRowUpdate
    '    'Dim dtSample As New DataTable
    '    'Dim drrow As DataRow
    '    ''Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    '    Try
    '        fillSampleAnalysis()

    '        'dtSample = dtAnalysis.Clone
    '        'For Each drrow In dtAnalysis.Rows
    '        '    dtSample.ImportRow(drrow)
    '        'Next
    '        'pClosure.SamplesTable = dtSample
    '        'For Each drRow In dtAnalysis.Rows
    '        '    dr = pClosure.SamplesTable.NewRow
    '        '    dr.Item("SAMPLE_ID") = drRow.Item("SAMPLE_ID")
    '        '    dr.Item("CLOSURE_ID") = drRow.Item("CLOSURE_ID")
    '        '    dr.Item("Sample #") = drRow.Item("Sample #")
    '        '    dr.Item("Analysis Type") = drRow.Item("Analysis Type")
    '        '    dr.Item("Analysis Level") = drRow.Item("Analysis Level")
    '        '    dr.Item("Sample Media") = drRow.Item("Sample Media")
    '        '    dr.Item("Sample Location") = drRow.Item("Sample Location")
    '        '    dr.Item("Sample Value") = drRow.Item("Sample Value")
    '        '    dr.Item("Sample Units") = drRow.Item("Sample Units")
    '        '    dr.Item("Sample Constituent") = drRow.Item("Sample Constituent")
    '        '    dr.Item("DELETED") = drRow.Item("DELETED")

    '        '    pClosure.SamplesTable.Rows.Add(dr)
    '        'Next



    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub dGridAnalysis_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles dGridAnalysis.BeforeRowsDeleted
        Dim drow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            e.DisplayPromptMsg = False
            Dim results As MsgBoxResult = MsgBox("You have selected " + e.Rows.Length.ToString + "row(s) for deletion." + _
                                                    "Do you want to continue", MsgBoxStyle.YesNo, "Delete Row(s)")
            If results = MsgBoxResult.Yes Then
                For Each drow In e.Rows
                    drow.Cells("DELETED").Value = True
                    SaveSampleTableRow(drow)
                Next
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dGridAnalysis_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dGridAnalysis.AfterRowsDeleted
        Try
            ' get the table from db
            RefreshAnalysisGrid(True)
            pClosure.SamplesTableOriginal = pClosure.SamplesTable
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dGridAnalysis_BeforeRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs) Handles dGridAnalysis.BeforeRowInsert
        Try
            If pClosure.ID <= 0 Then
                If MsgBox("Please Save Closure event before entering Analysis values" + vbCrLf + _
                        "Do you want to save Closure event?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
                    If pClosure.ID <= 0 Then
                        pClosure.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pClosure.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    If Not pClosure.Save(CType(UIUtilsGen.ModuleID.Closure, Integer), MusterContainer.AppUser.UserKey, returnVal) Then
                        e.Cancel = True
                    End If

                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                Else
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dGridAnalysis_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dGridAnalysis.AfterRowInsert
        Try
            e.Row.Cells("SAMPLE_ID").Value = 0
            e.Row.Cells("CLOSURE_ID").Value = pClosure.ID
            e.Row.Cells("DELETED").Value = False
            e.Row.Cells("Analysis Level").ValueList = New Infragistics.Win.ValueList
            e.Row.Cells("Sample Location").ValueList = New Infragistics.Win.ValueList

            dGridAnalysis.Focus()
            'e.Row.Cells("Sample #").Activate()

            dGridAnalysis.ActiveCell = dGridAnalysis.ActiveRow.Cells("Sample #")
            dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ToggleCellSel, False, False)
            dGridAnalysis.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
            ' to enable save button
            'setClosureSaveCancel(True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dGridAnalysis_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles dGridAnalysis.BeforeRowUpdate
        Try
            ' save changes to db
            If e.Row.Cells("Sample #").Value Is DBNull.Value Then
                MsgBox("Sample # is required")
                Exit Sub
            End If

            If e.Row.Cells("Analysis Level").Value Is DBNull.Value Then
                MsgBox("Analysis level is required")
                Exit Sub
            End If

            'Added by hcao on Oct. 25, 2007. 
            'Give warning message if ADL is selected at the Analysis Level.
            If e.Row.Cells("Analysis Level").Value = 884 Then
                If (MsgBox("You have selected 'ADL' in the Analysis Level. Do you want to continue saving it?", MsgBoxStyle.YesNo) <> MsgBoxResult.Yes) Then
                    Exit Sub
                End If
            End If

            If e.Row.Cells("Analysis Level").Value = 890 Then
                If (MsgBox("You have selected 'AAL' in the Analysis Level. Do you want to continue saving it?", MsgBoxStyle.YesNo) <> MsgBoxResult.Yes) Then
                    Exit Sub
                End If
            End If

            SaveSampleTableRow(e.Row)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dGridAnalysis_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles dGridAnalysis.AfterRowUpdate
        Try
            ' get the table from db
            RefreshAnalysisGrid(True)
            pClosure.SamplesTableOriginal = pClosure.SamplesTable
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dGridAnalysis_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles dGridAnalysis.CellChange
        Try
            If e.Cell.Row.Cells("Sample #").Text = String.Empty Or e.Cell.Row.Cells("Sample #").Text Is DBNull.Value Then
                MsgBox("Please enter Sample # first")
                e.Cell.CancelUpdate()
                Exit Try
            End If
            If "Sample #".Equals(e.Cell.Column.Key) Or _
                "Sample Constituent".Equals(e.Cell.Column.Key) Then
                e.Cell.Value = e.Cell.Text
                'ElseIf "Sample Value".Equals(e.Cell.Column.Key) Then
                ' #836
                ' no need to set value as value will be set to the col when focus is out of cell/row
                'If CType(e.Cell.Text, Single) < 0 Or CType(e.Cell.Text, Single) > 9999 Then
                '    MsgBox("Invalid value" + vbCrLf + "Value must be from 1 to 9999")
                '    e.Cell.CancelUpdate()
                'Else
                '    e.Cell.Value = e.Cell.Text
                'End If
            ElseIf "Analysis Type".Equals(e.Cell.Column.Key) Then
                e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                PopulateAnalysisLevel(e.Cell.Row)
            ElseIf "Analysis Level".Equals(e.Cell.Column.Key) Then
                e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                PopulateSampleLocation(e.Cell.Row)
            ElseIf "Sample Media".Equals(e.Cell.Column.Key) Or _
                "Sample Location".Equals(e.Cell.Column.Key) Or _
                "Sample Units".Equals(e.Cell.Column.Key) Then
                e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region
#End Region
#End Region
#Region "Comments"
    Private Sub CommentsMaintenance(Optional ByVal sender As System.Object = Nothing, Optional ByVal e As System.EventArgs = Nothing, Optional ByVal bolSetCounts As Boolean = False, Optional ByVal resetBtnColor As Boolean = False)
        Dim SC As ShowComments
        Dim nEntityType As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityName As String = String.Empty
        Dim oComments As MUSTER.BusinessLogic.pComments
        Dim bolEnableShowAllModules As Boolean = True
        Dim nCommentsCount As Integer = 0
        Try
            Select Case Me.tbCntrlClosure.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    strEntityName = "Owner : " + CStr(pOwn.ID) + " " + Me.txtOwnerName.Text
                    oComments = pOwn.Comments
                    nEntityID = pOwn.ID
                    nEntityType = UIUtilsGen.EntityTypes.Owner
                Case tbPageFacilityDetail.Name
                    strEntityName = "Facility : " + CStr(pOwn.Facilities.ID) + " " + pOwn.Facilities.Name
                    oComments = pOwn.Facilities.Comments
                    nEntityID = pOwn.Facilities.ID
                    nEntityType = UIUtilsGen.EntityTypes.Facility
                Case tbPageClosures.Name
                    strEntityName = "Facility : " + CStr(pOwn.Facilities.ID) + " " + pOwn.Facilities.Name + ", Closure Event : " + CStr(pClosure.FacilitySequence)
                    oComments = pClosure.Comments
                    nEntityID = pClosure.ID
                    nEntityType = UIUtilsGen.EntityTypes.ClosureEvent
                    bolEnableShowAllModules = False
                Case Else
                    Exit Sub
            End Select
            If Not resetBtnColor Then
                SC = New ShowComments(nEntityID, nEntityType, IIf(bolSetCounts, "", "Closure"), strEntityName, oComments, Me.Text, , bolEnableShowAllModules)
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
            ElseIf nEntityType = UIUtilsGen.EntityTypes.ClosureEvent Then
                If nCommentsCount > 0 Then
                    btnClosureComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnClosureComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
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
        Dim MyFrm As MusterContainer
        MyFrm = Me.MdiParent
        'MyFrm.FlagsChanged(entityID, entityType, [Module], ParentFormText)
        ' if in closure tab, pass closure id
        If tbCntrlClosure.SelectedTab.Name = tbPageClosures.Name Then
            If pClosure.ID > 0 Then
                MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text, pClosure.ID, UIUtilsGen.EntityTypes.ClosureEvent)
            Else
                MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text)
            End If
        Else
            MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text)
        End If
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        Dim mc As MusterContainer = Me.MdiParent
        mc.RefreshCalendarInfo()
        mc.LoadDueToMeCalendar()
        mc.LoadToDoCalendar()
    End Sub
    Private Sub FlagMaintenance(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ownID, facID As Integer
            Select Case Me.tbCntrlClosure.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    SF = New ShowFlags(pOwn.ID, UIUtilsGen.EntityTypes.Owner, "Closure")
                Case tbPageFacilityDetail.Name
                    SF = New ShowFlags(pOwn.Facilities.ID, UIUtilsGen.EntityTypes.Facility, "Closure")
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
        'MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, oEntity.ID)
    End Sub
    Private Sub btnFacFlags_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacFlags.Click
        FlagMaintenance(sender, e)
        'Dim MyFrm As MusterContainer
        'MyFrm = Me.MdiParent
        'oEntity.GetEntity("Facility")
        'MyFrm.FlagsChanged(Me.lblFacilityIDValue.Text, oEntity.ID)
    End Sub
#End Region
#End Region
#Region "External Event Handlers"
#Region "Owner"
    Public Sub evtOwnerErr(ByVal ErrMsg As String) Handles pOwn.evtOwnerErr
        If Not bolDisplayErrmessage Then
            'If nErrMessage > 0 Then
            bolValidateSuccess = True
            Exit Sub
        End If
        If ErrMsg <> String.Empty And MsgBox(ErrMsg) = MsgBoxResult.OK Then
            ErrMsg = String.Empty
        End If
        bolDisplayErrmessage = False
        bolValidateSuccess = False
    End Sub
    Public Sub PersonaChanged(ByVal bolState As Boolean) Handles pOwn.evtPersonaChanged
        SetPersonaSaveCancel(bolState)
    End Sub
    Public Sub PersonasChanged(ByVal bolstate As Boolean) Handles pOwn.evtPersonasChanged
        SetPersonaSaveCancel(bolstate)
    End Sub
    Public Sub OwnersChanged(ByVal bolstate As Boolean) Handles pOwn.evtOwnersChanged
        SetOwnerSaveCancel(bolstate)
    End Sub
    Public Sub OwnerChanged(ByVal bolstate As Boolean) Handles pOwn.evtOwnerChanged
        SetOwnerSaveCancel(bolstate)
    End Sub
#End Region
#Region "Facility"
    Public Sub FacilityChanged(ByVal bolstate As Boolean) Handles pOwn.evtFacilityChanged
        SetFacilitySaveCancel(bolstate)
    End Sub
    Public Sub FacilitiesChanged(ByVal bolstate As Boolean) Handles pOwn.evtFacilitiesChanged
        SetFacilitySaveCancel(bolstate)
    End Sub
    Public Sub ValidationErrors(ByVal FacID As Integer, ByVal MsgStr As String) Handles pOwn.evtValidationErr
        If Not Me.bolDisplayErrmessage Then
            Me.bolDisplayErrmessage = True
            Exit Sub
        End If
        pOwn.Facilities.ID = FacID
        If MsgStr <> String.Empty And MsgBox(MsgStr + " on FacilityID" + FacID.ToString) = MsgBoxResult.OK Then
            MsgStr = String.Empty
        End If
        bolValidateSuccess = False
        Me.bolDisplayErrmessage = False
    End Sub
#End Region
#Region "Closure"
    Private Sub pClosure_evtClosureEventErr(ByVal MsgStr As String) Handles pClosure.evtClosureEventErr
        'If Not bolDisplayErrmessage Then
        If nErrMessage > 0 Then
            bolDisplayErrmessage = True
            nErrMessage = 0
            Exit Sub
        End If
        If MsgStr <> String.Empty And MsgBox(MsgStr) = MsgBoxResult.OK Then
            MsgStr = String.Empty
            nErrMessage = 1
            bolDisplayErrmessage = False
        End If
        'bolDisplayErrmessage = False
        bolValidateSuccess = False
    End Sub
    Private Sub pClosure_evtClosureLetter(ByVal letterType As MUSTER.BusinessLogic.pClosureEvent.LetterType) Handles pClosure.evtClosureLetter
        Dim oLetter As New Reg_Letters
        Dim oUser As New MUSTER.BusinessLogic.pUser
        Dim oUserInfo As New MUSTER.Info.UserInfo
        Dim oUserInfoClosureHead As New MUSTER.Info.UserInfo
        Try
            oUserInfo = oUser.RetrievePMHead
            oUserInfoClosureHead = oUser.RetrieveClosureHead
            If slLetterCount Is Nothing Then
                ' initialize sorted list to maintain the letter count to prevent creating multiple letters
                InitializeLetterCount()
            End If
            Select Case letterType
                Case letterType.InfoNeeded
                    If CType(slLetterCount.Item(letterType.InfoNeeded), Integer) > 0 Then
                        Exit Sub
                    ElseIf lblFacilityIDValue.Text = String.Empty Then
                        Exit Sub
                    ElseIf pClosure.FacilityID <> CInt(lblFacilityIDValue.Text) Then
                        Exit Sub
                    End If
                    Dim colInfoNeeded As ArrayList = New ArrayList
                    Dim slchkCol As New SortedList
                    Dim sldtPickCol As New SortedList
                    For Each ctrl As Control In pnlChecklistDetails.Controls
                        If ctrl.GetType.ToString.ToLower = "System.Windows.Forms.CheckBox".ToLower Then
                            If ctrl.Name.StartsWith("chkClosures") Then
                                slchkCol.Add(CType(ctrl.Tag, Integer), ctrl)
                            End If
                        ElseIf ctrl.GetType.ToString.ToLower = "System.Windows.Forms.DateTimePicker".ToLower Then
                            If ctrl.Name.StartsWith("dtPickClosures") Then
                                sldtPickCol.Add(CType(ctrl.Tag, Integer), ctrl)
                            End If
                        End If
                    Next

                    For Each htEntry As DictionaryEntry In pClosure.HashTableBoolCheckList
                        If CType(htEntry.Value, Boolean) = True Then
                            If Date.Compare(CType(pClosure.HashTableDateCheckList.Item(htEntry.Key), Date), CDate("01/01/0001")) = 0 Then
                                Dim chkBox As CheckBox
                                chkBox = slchkCol.GetByIndex(slchkCol.IndexOfKey(htEntry.Key))
                                colInfoNeeded.Add(chkBox.Text.Trim)
                            End If
                        End If
                    Next
                    If bolFromNOI Then
                        oLetter.GenerateClosureInfoNeededLetter(pClosure.FacilityID, "Information Needed Letter", "InfoNeeded_Letter", "Information Needed Letter for Closure", "InfoNeeded.doc", pOwn, pClosure, colInfoNeeded, txtNOILicensee.Text)
                    Else
                        oLetter.GenerateClosureInfoNeededLetter(pClosure.FacilityID, "Information Needed Letter", "InfoNeeded_Letter", "Information Needed Letter for Closure", "InfoNeeded.doc", pOwn, pClosure, colInfoNeeded, txtLicensee.Text)
                    End If
                    slLetterCount.Item(letterType.InfoNeeded) = 1
                    'oLetter.GenerateClosureLetter(pClosure.FacilityID, "Information Needed Letter", "InfoNeeded_Letter", "Information Needed Letter for Closure", "USTLetterInfoNeeded_Template1.doc", pClosure.DueDate, pOwn, strAnalysisType, oUserInfo.ID)
                Case letterType.RFGApproval
                    If CType(slLetterCount.Item(letterType.RFGApproval), Integer) > 0 Then
                        Exit Sub
                    ElseIf lblFacilityIDValue.Text = String.Empty Then
                        Exit Sub
                    ElseIf pClosure.FacilityID <> CInt(lblFacilityIDValue.Text) Then
                        Exit Sub
                    End If
                    oLetter.GenerateClosureLetter(pClosure.FacilityID, "RFG Approval Letter", "RFG_Approval_Letter", "Removal From Ground Approval Letter for Closure", "Approval_RFG_Template.doc", pClosure.ScheduledDate, pOwn, strAnalysisType, oUserInfo.Name, pClosure.ID, txtNOILicensee.Text, pClosure.ID, pClosure.FacilitySequence, UIUtilsGen.EntityTypes.ClosureEvent, strMedia)
                    slLetterCount.Item(letterType.RFGApproval) = 1
                Case letterType.CIPApproval
                    If CType(slLetterCount.Item(letterType.CIPApproval), Integer) > 0 Then
                        Exit Sub
                    ElseIf lblFacilityIDValue.Text = String.Empty Then
                        Exit Sub
                    ElseIf pClosure.FacilityID <> CInt(lblFacilityIDValue.Text) Then
                        Exit Sub
                    End If
                    oLetter.GenerateClosureLetter(pClosure.FacilityID, "CIP Approval Letter", "CIP_Approval_Letter", "Closure In Place Approval Letter for Closure", "Approval_CIP_Template.doc", pClosure.DueDate, pOwn, strAnalysisType, oUserInfo.Name, pClosure.ID, txtNOILicensee.Text, pClosure.ID, pClosure.FacilitySequence, UIUtilsGen.EntityTypes.ClosureEvent, strMedia)
                    slLetterCount.Item(letterType.CIPApproval) = 1
                Case letterType.CIPDisApproval
                    If CType(slLetterCount.Item(letterType.CIPDisApproval), Integer) > 0 Then
                        Exit Sub
                    ElseIf lblFacilityIDValue.Text = String.Empty Then
                        Exit Sub
                    ElseIf pClosure.FacilityID <> CInt(lblFacilityIDValue.Text) Then
                        Exit Sub
                    End If
                    oLetter.GenerateClosureLetter(pClosure.FacilityID, "CIP DisApproval Letter", "CIP_DisApproval_Letter", "Closure In Place DisApproval Letter for Closure", "DisApproval_CIP_Template.Doc", pClosure.DueDate, pOwn, strAnalysisType, oUserInfo.Name, pClosure.ID, txtNOILicensee.Text, pClosure.ID, pClosure.FacilitySequence, UIUtilsGen.EntityTypes.ClosureEvent, strMedia)
                    slLetterCount.Item(letterType.CIPDisApproval) = 1

                Case letterType.CIPNFA
                    If CType(slLetterCount.Item(letterType.CIPNFA), Integer) > 0 Then
                        Exit Sub
                    ElseIf lblFacilityIDValue.Text = String.Empty Then
                        Exit Sub
                    ElseIf pClosure.FacilityID <> CInt(lblFacilityIDValue.Text) Then
                        Exit Sub
                    End If
                    oLetter.GenerateClosureLetter(pClosure.FacilityID, "CIP NFA Letter", "CIPNFA_Letter", "CIP No Further Activity Letter for Closure", IIf(pClosure.CRDateLastUsed < "12/22/1988", "Pre88_NFA_template.doc", "NFA_Template.doc"), pClosure.DueDate, pOwn, strAnalysisType, oUserInfo.Name, pClosure.ID, , pClosure.ID, pClosure.FacilitySequence, UIUtilsGen.EntityTypes.ClosureEvent, strMedia)
                    slLetterCount.Item(letterType.CIPNFA) = 1
                Case letterType.RFGNFA
                    If CType(slLetterCount.Item(letterType.RFGNFA), Integer) > 0 Then
                        Exit Sub
                    ElseIf lblFacilityIDValue.Text = String.Empty Then
                        Exit Sub
                    ElseIf pClosure.FacilityID <> CInt(lblFacilityIDValue.Text) Then
                        Exit Sub
                    End If
                    oLetter.GenerateClosureLetter(pClosure.FacilityID, "RFG NFA Letter", "RFGNFA_Letter", "RFG No Further Activity Letter for Closure", IIf(pClosure.CRDateLastUsed < "12/22/1988", "Pre88_NFA_template.doc", "NFA_Template.doc"), pClosure.DueDate, pOwn, strAnalysisType, oUserInfo.Name, pClosure.ID, , pClosure.ID, pClosure.FacilitySequence, UIUtilsGen.EntityTypes.ClosureEvent, strMedia)
                    slLetterCount.Item(letterType.RFGNFA) = 1
                Case letterType.SampleResultMemo
                    If CType(slLetterCount.Item(letterType.SampleResultMemo), Integer) > 0 Then
                        Exit Sub
                    ElseIf lblFacilityIDValue.Text = String.Empty Then
                        Exit Sub
                    ElseIf pClosure.FacilityID <> CInt(lblFacilityIDValue.Text) Then
                        Exit Sub
                    End If
                    oLetter.GenerateSampleDemo(pClosure.FacilityID, "Sample Memo Letter", "Sample_Memo_Letter", "Sample Memo Letter for Closure", "SampleResultDemo.doc", pClosure.DueDate, pOwn, oUserInfo.Name, oUserInfoClosureHead.Name, pClosure.SampleResultMemo, pClosure.ID, , pClosure.ID, pClosure.FacilitySequence, UIUtilsGen.EntityTypes.ClosureEvent)
                    slLetterCount.Item(letterType.SampleResultMemo) = 1
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub pClosure_evtClosureEventInfoChanged(ByVal bolValue As Boolean)
        setClosureSaveCancel(bolValue)
    End Sub
    Private Sub pClosure_evtClosureCRFillMaterial(ByVal bol As Boolean) Handles pClosure.evtClosureCRFillMaterial
        If bol Then
            bolLoading = True
            Me.cmbFillMaterial.Enabled = dtPickReceived.Enabled
            Me.cmbClosureReportFillMaterial.Enabled = True
            PopulateClosureFillMaterial()
            bolLoading = False
        Else
            Me.cmbFillMaterial.Enabled = False
            Me.cmbClosureReportFillMaterial.Enabled = False
        End If
    End Sub
    Private Sub pClosure_evtClosureNOIReceived(ByVal bolEnable As Boolean) Handles pClosure.evtClosureNOIReceived
        Dim bolLoadingLocal As Boolean = bolLoading
        If Not bolLoadingLocal Then bolLoading = True
        dtPickReceived.Enabled = bolEnable
        dtPickScheduledDate.Enabled = bolEnable
        txtNOILicensee.Enabled = bolEnable
        txtNOICompany.Enabled = bolEnable
        lblNOISearchCompany.Enabled = bolEnable
        UIUtilsGen.SetDatePickerValue(dtPickReceived, pClosure.NOI_Rcv_Date)
        UIUtilsGen.SetDatePickerValue(dtPickScheduledDate, pClosure.ScheduledDate)
        If Not bolEnable Then
            btnProcessNOI.Enabled = bolEnable
            btnProcessNOIEnvelopes.Enabled = bolEnable
            lblNoticeOfInterestDisplay.Text = "+"
            PnlNoticeOfInterestDetails.Visible = False
        Else
            If pClosure.NOIProcessed Then
                btnProcessNOI.Enabled = Not bolEnable
                btnProcessNOIEnvelopes.Enabled = Not bolEnable
            Else
                btnProcessNOI.Enabled = bolEnable
                btnProcessNOIEnvelopes.Enabled = bolEnable
            End If
            lblNoticeOfInterestDisplay.Text = "-"
            PnlNoticeOfInterestDetails.Visible = True
            cmbFillMaterial.Enabled = cmbClosureReportFillMaterial.Enabled
            If cmbFillMaterial.Enabled Then PopulateClosureFillMaterial()
            If pClosure.FillMaterial <> 0 Then UIUtilsGen.SetComboboxItemByValue(cmbFillMaterial, pClosure.FillMaterial)
        End If
        If Not bolLoadingLocal Then bolLoading = False
    End Sub
    Private Sub pClosure_FlagsChanged(ByVal entityID As Integer, ByVal entityType As Integer, ByVal eventID As Integer, ByVal eventType As Integer) Handles pClosure.FlagsChanged
        Dim mc As MusterContainer = Me.MdiParent
        If Not mc Is Nothing Then
            mc.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Closure", Me.Text, eventID, eventType)
        End If
    End Sub
#End Region
#End Region
#Region "Contacts"
#Region "Owner Contacts"
    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        If bolLoading Then Exit Sub
        Try
            Select Case TabControl1.SelectedTab.Name
                'Case "tbPrevFacs".ToUpper
                '    TabControl1.SelectedTab = tbPageOwnerFacilities
            Case tbPageOwnerFacilities.Name
                    'If tbCtrlOwner.Contains(Me.tbPrevFacs) Then
                    '    Me.tbCtrlOwner.TabPages.RemoveAt(1)
                    'End If
                Case tbPageOwnerContactList.Name
                    If Me.lblOwnerIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, Integer.Parse(Me.lblOwnerIDValue.Text), UIUtilsGen.EntityTypes.Owner)
                    End If
                    'Dim dsOwnerContacts As DataSet
                    'dsOwnerContacts = MusterContainer.pConStruct.GetAll()
                    'dsOwnerContacts.Tables(0).DefaultView.RowFilter = "MODULEID = 891 And ENTITYID = " + pOwn.ID.ToString
                    'dsContacts.Tables(0).DefaultView.Sort = "CONTACT_NAME ASC"
                    'ugOwnerContacts.DataSource = dsOwnerContacts.Tables(0).DefaultView
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("Parent_Contact").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("Child_Contact").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("ContactAssocID").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("ModuleID").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("ContactType").Hidden = True
                Case tbPageOwnerDocuments.Name
                    UCOwnerDocuments.LoadDocumentsGrid(Integer.Parse(Me.lblOwnerIDValue.Text), UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Closure)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerAddSearchContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAddSearchContact.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                objCntSearch = New ContactSearch(pOwn.ID, 9, "Closure", pConStruct)
            Else
                objCntSearch = New ContactSearch(pOwn.Facility.ID, 6, "Closure", pConStruct)
            End If
            'objCntSearch.Show()
            objCntSearch.ShowDialog()

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
            If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                AssociateContact(ugOwnerContacts, pOwn.ID, 9)
            Else
                AssociateContact(ugOwnerContacts, pOwn.Facility.ID, 6)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub btnOwnerDeleteContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerDeleteContact.Click
        Try
            If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                DeleteContact(ugOwnerContacts, pOwn.ID)
            Else
                DeleteContact(ugOwnerContacts, pOwn.Facility.ID)
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
    Dim strFilterString As String = String.Empty
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
    Private Sub chkOwnerShowRelatedContacts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowRelatedContacts.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFilter()
            'End If
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
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    
            If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                nEntityID = pOwn.ID
                strMode = UCase("owner")
                nEntityType = 9
            Else
                nEntityID = pOwn.Facility.ID
                nEntityType = 6
            End If

            If chkOwnerShowActiveOnly.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If

            If chkOwnerShowContactsforAllModules.Checked Then

                ' User has the ability to view the contacts associated for the entity in other modules
                If strMode <> "OWNER" Then
                    Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facility.ID.ToString)
                    strEntityAssocIDs = strFilterForAllModules
                    nModuleID = 0
                End If
            Else
                nModuleID = 891
            End If
            If chkOwnerShowRelatedContacts.Checked Then
                strEntities = strFacilityIdTags
                nRelatedEntityType = 6
            End If

            UIUtilsGen.LoadContacts(ugOwnerContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, strEntityAssocIDs, nRelatedEntityType)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'strFilterString = String.Empty
            'Dim strEntityID As String
            'If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
            '    strEntityID = pOwn.ID.ToString
            '    strMode = UCase("owner")
            'Else
            '    strEntityID = pOwn.Facility.ID.ToString
            'End If

            'If chkOwnerShowActiveOnly.Checked Then
            '    strFilterString = "ACTIVE = 1"
            'Else
            '    strFilterString = ""
            'End If

            'If chkOwnerShowContactsforAllModules.Checked Then

            '    ' User has the ability to view the contacts associated for the entity in other modules
            '    If strMode <> "OWNER" Then
            '        Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facility.ID.ToString)
            '        If strFilterString = String.Empty Then
            '            'strFilterString += "ENTITYID = " + strEntityID + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '            strFilterString += "ENTITYID = " + strEntityID + IIf(Not strFilterForAllModules = String.Empty, " OR " + " entityassocid in (" + strFilterForAllModules + ")", "")
            '        Else
            '            'strFilterString += " AND ENTITYID = " + strEntityID + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '            strFilterString += " AND ENTITYID = " + strEntityID + IIf(Not strFilterForAllModules = String.Empty, " OR " + " entityassocid in (" + strFilterForAllModules + ")", "")
            '        End If
            '    Else
            '        If strFilterString = String.Empty Then
            '            strFilterString += "ENTITYID = " + strEntityID
            '        Else
            '            strFilterString += " AND ENTITYID = " + strEntityID
            '        End If
            '    End If
            'Else
            '    If strFilterString = String.Empty Then
            '        strFilterString += " MODULEID = 891 And ENTITYID = " + strEntityID
            '    Else
            '        strFilterString += " AND MODULEID = 891 And ENTITYID = " + strEntityID
            '    End If
            'End If
            'If chkOwnerShowRelatedContacts.Checked Then
            '    'strFilterString = " (ENTITYID = " + strEntityID + IIf(Not strFacilityIdTags = String.Empty, " OR ENTITYID in (" + strFacilityIdTags + ")", "") + ")"
            '    strFilterString += IIf(Not strFacilityIdTags = String.Empty, " OR " + " ENTITYID in (" + strFacilityIdTags + ")", "")
            'Else
            '    strFilterString += ""
            'End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugOwnerContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "ClosureEvent Contacts"
    Private Sub chkClosureShowContactsForAllModules_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkClosureShowContactsForAllModules.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            ClosureSetFilter()
            'End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub chkClosureShowRelatedContacts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkClosureShowRelatedContacts.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            ClosureSetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkClosureShowActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkClosureShowActive.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            ClosureSetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClosureContactAddorSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClosureContactAddorSearch.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            objCntSearch = New ContactSearch(pClosure.ID, 22, "Closure", pConStruct)
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClosureContactModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClosureContactModify.Click
        Try
            ModifyContact(ugClosureContacts)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClosureContactDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClosureContactDelete.Click
        Try
            DeleteContact(ugClosureContacts, pClosure.ID)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClosureContactAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClosureContactAssociate.Click
        Try
            AssociateContact(ugClosureContacts, pClosure.ID, 22)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub btnClosureContactClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub
    Private Sub ClosureSetFilter()
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

            'strFilterString = String.Empty
            'Dim strEntityID As String

            'strEntityID = pClosure.ID.ToString

            If chkClosureShowActive.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If

            If chkClosureShowContactsForAllModules.Checked Then
                Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facility.ID.ToString)
                strEntityAssocIDs = strFilterForAllModules
                nEntityID = pOwn.Facility.ID
                nEntityType = 6
                nModuleID = 0
                'If strFilterString = "(" Then
                '    strFilterString += "ENTITYID = " + pOwn.Facility.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
                'Else
                '    strFilterString += "AND ENTITYID = " + pOwn.Facility.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
                'End If
            Else
                nEntityType = 22
                nEntityID = pClosure.ID
                nModuleID = 891
            End If

            If chkClosureShowRelatedContacts.Checked Then
                strEntities = strClosureEventIdTags
                nRelatedEntityType = 22
                'strFilterString += " OR " + IIf(Not strClosureEventIdTags = String.Empty, " ENTITYID in (" + strClosureEventIdTags + "))", "")
            End If
            UIUtilsGen.LoadContacts(ugClosureContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, strEntityAssocIDs, nRelatedEntityType)

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'strFilterString = String.Empty
            'Dim strEntityID As String

            'strEntityID = pClosure.ID.ToString

            'If chkClosureShowActive.Checked Then
            '    strFilterString = "(ACTIVE = 1"
            'Else
            '    strFilterString = "("
            'End If

            'If chkClosureShowContactsForAllModules.Checked Then
            '    'If strFilterString = "(" Then
            '    '    strFilterString += "ENTITYID = " + pOwn.Facility.ID.ToString + " OR ENTITYID = " + pClosure.ID.ToString
            '    'Else
            '    '    strFilterString += " AND ENTITYID = " + pOwn.Facility.ID.ToString + " OR ENTITYID = " + pClosure.ID.ToString
            '    'End If
            '    Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facility.ID.ToString)
            '    If strFilterString = "(" Then
            '        strFilterString += "ENTITYID = " + pOwn.Facility.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '    Else
            '        strFilterString += "AND ENTITYID = " + pOwn.Facility.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '    End If
            'Else
            '    If strFilterString = "(" Then
            '        strFilterString += " MODULEID = 891 And ENTITYID = " + strEntityID
            '    Else
            '        strFilterString += " AND MODULEID = 891 And ENTITYID = " + strEntityID
            '    End If
            'End If

            'If chkClosureShowRelatedContacts.Checked Then
            '    strFilterString += " OR " + IIf(Not strClosureEventIdTags = String.Empty, " ENTITYID in (" + strClosureEventIdTags + "))", "")
            'Else
            '    strFilterString += ")"
            'End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugClosureContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Common Functions"
    Private Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer)

        Try

            UIUtilsGen.LoadContacts(ugGrid, EntityID, EntityType, pConStruct, 891)

            If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Or tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEFACILITYDETAIL" Then
                Me.chkOwnerShowActiveOnly.Checked = False
                Me.chkOwnerShowActiveOnly.Checked = True
            ElseIf tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGECLOSURES" Then
                Me.chkClosureShowActive.Checked = False
                Me.chkClosureShowActive.Checked = True
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Function ModifyContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try

            If UIUtilsGen.ModifyContact(ugGrid, 891, pConStruct) Then
                Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer)
        Try

            If UIUtilsGen.AssociateContact(ugGrid, nEntityID, nEntityType, 891, pConStruct) Then
                Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer)
        Try

            If UIUtilsGen.DeleteContact(ugGrid, nEntityID, 891, pConStruct) Then
                Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

#End Region
#Region "Close Events"
    Private Sub Search_ContactAdded() Handles objCntSearch.ContactAdded
        If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGECLOSURES" Then
            LoadContacts(ugClosureContacts, pClosure.ID, 22)
            chkClosureShowContactsForAllModules.Checked = False
        Else
            If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                LoadContacts(ugOwnerContacts, pOwn.ID, 9)
            Else
                LoadContacts(ugOwnerContacts, pOwn.Facility.ID, 6)
            End If
            chkClosureShowContactsForAllModules.Checked = False
        End If
    End Sub
    Private Sub Contact_ContactAdded()
        If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGECLOSURES" Then
            LoadContacts(ugClosureContacts, pClosure.ID, 22)
            chkClosureShowContactsForAllModules.Checked = False
        Else
            If tbCntrlClosure.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                LoadContacts(ugOwnerContacts, pOwn.ID, 9)
            Else
                LoadContacts(ugOwnerContacts, pOwn.Facility.ID, 6)
            End If
            chkClosureShowContactsForAllModules.Checked = False
        End If
    End Sub

    Private Sub objCntSearch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objCntSearch.Closing
        If Not objCntSearch Is Nothing Then
            objCntSearch = Nothing
        End If
    End Sub
#End Region


#End Region
#Region "Envelopes and Labels"
    Private Sub btnClosureOwnerEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClosureOwnerEnvelopes.Click
        Dim arrAddress(4) As String
        Try

            arrAddress(0) = pOwn.Addresses.AddressLine1
            arrAddress(1) = pOwn.Addresses.AddressLine2
            arrAddress(2) = pOwn.Addresses.City
            arrAddress(3) = pOwn.Addresses.State
            arrAddress(4) = pOwn.Addresses.Zip
            'strAddress = pOwn.Addresses.AddressLine1 + "," + pOwn.Addresses.AddressLine2 + "," + pOwn.Addresses.City + "," + pOwn.Addresses.State + "," + pOwn.Addresses.Zip


            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then

            End If
            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then

                Dim dsContactsLocal = pConStruct.GetFilteredContacts(pOwn.ID, 612)

                Dim strContactName As String

                If Not dsContactsLocal Is Nothing AndAlso dsContactsLocal.Tables.Count > 0 AndAlso dsContactsLocal.Tables(0).Rows.Count > 0 Then
                    For Each contactRow As DataRow In dsContactsLocal.Tables(0).Rows
                        If contactRow("Type") = "Registration Representative" Then
                            strContactName = contactRow("CONTACT_name")
                            Exit For
                        End If

                    Next
                Else

                    If Not pOwn.Persona.LastName.Length > 0 Then
                        strContactName = String.Format("{0}{1}{2}{3}", pOwn.Persona.FirstName, IIf(pOwn.Persona.MiddleName.Length > 0, " ", ""), pOwn.Persona.MiddleName, pOwn.Persona.LastName)
                    Else
                        strContactName = pOwn.Persona.Company
                    End If

                End If

                UIUtilsGen.CreateEnvelopes(Me.txtOwnerName.Text, arrAddress, "CLO", pOwn.ID, strContactName)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClosureOwnerLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClosureOwnerLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Addresses.AddressLine1
            arrAddress(1) = pOwn.Addresses.AddressLine2
            arrAddress(2) = pOwn.Addresses.City
            arrAddress(3) = pOwn.Addresses.State
            arrAddress(4) = pOwn.Addresses.Zip
            'strAddress = pOwn.Addresses.AddressLine1 + "," + pOwn.Addresses.AddressLine2 + "," + pOwn.Addresses.City + "," + pOwn.Addresses.State + "," + pOwn.Addresses.Zip
            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtOwnerName.Text, arrAddress, "CLO", pOwn.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClosureFacEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClosureFacEnvelopes.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(Me.txtFacilityName.Text, arrAddress, "CLO", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClosureFacLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClosureFacLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtFacilityName.Text, arrAddress, "CLO", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnProcessNOIEnvelopes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcessNOIEnvelopes.Click
        Dim oLicAddInfo As MUSTER.Info.ComAddressInfo
        Dim oLicAdd As New MUSTER.BusinessLogic.pComAddress
        Dim arrAddress(4) As String
        Dim strName As String
        Try
            ' <Certified Contractor>
            Dim dsSearchResult As DataSet = pOwn.RunSQLQuery("SELECT COM_ADDRESS_ID FROM tblCOM_COMPANY_LICENSEE WHERE DELETED = 0 AND COMPANY_ID = " + pClosure.Company.ToString + " AND LICENSEE_ID = " + pClosure.CertContractor.ToString)
            For Each dr As DataRow In dsSearchResult.Tables(0).Rows
                If Not dr("COM_ADDRESS_ID") Is DBNull.Value Then
                    oLicAddInfo = oLicAdd.Retrieve(dr("COM_ADDRESS_ID"), False)
                    arrAddress(0) = oLicAddInfo.AddressLine1
                    arrAddress(1) = oLicAddInfo.AddressLine2
                    arrAddress(2) = oLicAddInfo.City
                    arrAddress(3) = oLicAddInfo.State
                    arrAddress(4) = oLicAddInfo.Zip
                    strName = txtNOILicensee.Text.Trim

                    If Not Me.txtNOICompany.Text Is Nothing AndAlso Me.txtNOICompany.Text.Length > 0 Then
                        strName = String.Format("{0}{1}{2}", strName, Chr(13), txtNOICompany.Text)
                    End If

                    UIUtilsGen.CreateEnvelopes(strName, arrAddress, "CLO", pClosure.ID)
                End If
            Next

            ' <CList> - CC List
            Dim dsContacts As DataSet = pConStruct.GetContactsByEntityAndModule(pClosure.ID, UIUtilsGen.ModuleID.Closure)
            If Not dsContacts Is Nothing Then
                If dsContacts.Tables.Count > 0 Then
                    If dsContacts.Tables(0).Rows.Count > 0 Then
                        For Each dr As DataRow In dsContacts.Tables(0).Rows
                            If Not dr.Item("CC_Info") Is System.DBNull.Value Then
                                If dr.Item("CC_Info") = "YES" Then

                                    arrAddress(0) = ""
                                    arrAddress(1) = ""
                                    arrAddress(2) = ""
                                    arrAddress(3) = ""
                                    arrAddress(4) = ""

                                    If Not dr.Item("DISPLAYAS") Is System.DBNull.Value Then
                                        If dr.Item("DISPLAYAS") <> String.Empty Then
                                            strName = dr.Item("DISPLAYAS")
                                        Else
                                            strName = dr.Item("Contact_Name")
                                        End If
                                    Else
                                        strName = dr.Item("Contact_Name")
                                    End If

                                    ' get company address
                                    If Not dr.Item("child_contact") Is DBNull.Value Then
                                        pConStruct.ContactDatum.Retrieve(dr.Item("child_contact"), False)
                                        If pConStruct.ContactDatum.ID = dr.Item("child_contact") Then
                                            arrAddress(0) = pConStruct.ContactDatum.AddressLine1
                                            arrAddress(1) = pConStruct.ContactDatum.AddressLine2
                                            arrAddress(2) = pConStruct.ContactDatum.City
                                            arrAddress(3) = pConStruct.ContactDatum.State
                                            arrAddress(4) = pConStruct.ContactDatum.ZipCode
                                        End If
                                    Else ' use licensee address
                                        If Not dr.Item("Address_One") Is DBNull.Value Then
                                            arrAddress(0) = dr.Item("Address_One")
                                        End If
                                        If Not dr.Item("Address_Two") Is DBNull.Value Then
                                            arrAddress(1) = dr.Item("Address_Two")
                                        End If
                                        If Not dr.Item("City") Is DBNull.Value Then
                                            arrAddress(2) = dr.Item("City")
                                        End If
                                        If Not dr.Item("State") Is DBNull.Value Then
                                            arrAddress(3) = dr.Item("State")
                                        End If
                                        If Not dr.Item("Zip") Is DBNull.Value Then
                                            arrAddress(4) = dr.Item("Zip")
                                        End If
                                    End If

                                    If arrAddress(0) <> String.Empty Then
                                        UIUtilsGen.CreateEnvelopes(strName, arrAddress, "CLO", pClosure.ID)
                                    End If

                                End If
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
    Private Sub lblLicenseeSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblLicenseeSearch.Click
        oCompanySearch = New CompanySearch
        oCompanySearch.ShowDialog()
    End Sub
    Private Sub LicenseeCompanyDetails(ByVal Licensee_id As Integer, ByVal company_id As Integer, ByVal Licensee_name As String, ByVal company_name As String) Handles oCompanySearch.LicenseeCompanyDetails
        Try
            If Not bolFromNOI Then
                txtLicensee.Text = Licensee_name
                txtClosureReportCompany.Text = company_name
                pClosure.CRCompany = company_id
                pClosure.CRCertContractor = Licensee_id
            Else
                txtNOILicensee.Text = Licensee_name
                txtNOICompany.Text = company_name
                pClosure.Company = company_id
                pClosure.CertContractor = Licensee_id
                If pClosure.CRCompany = 0 And pClosure.CRCertContractor = 0 Then
                    pClosure.CRCompany = company_id
                    pClosure.CRCertContractor = Licensee_id
                    txtLicensee.Text = Licensee_name
                    txtClosureReportCompany.Text = company_name
                End If
                bolFromNOI = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub lblNOISearchCompany_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblNOISearchCompany.Click
        bolFromNOI = True
        oCompanySearch = New CompanySearch
        oCompanySearch.ShowDialog()
    End Sub

    Private Sub tbCtrlFacClosureEvts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlFacClosureEvts.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            Select Case tbCtrlFacClosureEvts.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If Me.lblFacilityIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility)
                    End If
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Closure)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
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

    Private Sub ugClosureContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugClosureContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ugClosureContacts)
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
            If Me.btnSaveClosure.Enabled = False Then
                Me.btnSaveClosure.Enabled = True
                Me.btnClosureCancel.Enabled = True
            End If
        End If
    End Sub

End Class
