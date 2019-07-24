Public Class TransferOwnership
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.ShowComment.vb
    '   Provides the interface for transferring ownership of facilities.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1        JC      1/02/04    Changed AppUser.UserName to AppUser.ID to
    '                                  accomodate new use of pUser by application.
    '                                 
    '-------------------------------------------------------------------------------
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"
    Dim nOwnerID As Integer
    'Private oReg As InfoRepository.Registrations
    Private oRegGroupRight As New MUSTER.BusinessLogic.pRegistration
    Private oRegActinfo As MUSTER.Info.RegistrationActivityInfo
    Private pOwn As MUSTER.BusinessLogic.pOwner
    ' to handle the owner on the right side, so it does not interfere 
    ' with the current owner object
    Private pOwn2 As MUSTER.BusinessLogic.pOwner
    'Private frmRegistration As Registration
    Dim dtGlobal As DataTable
    Dim sortIndex As String = String.Empty

    Private bolFacTransferred As Boolean = False
    Dim bolLoading As Boolean = False
    Dim returnVal As String = String.Empty
#End Region

#Region " Windows Form Designer generated code "

    'Public Sub New(Optional ByRef frmMusterContainer As MusterContainer = Nothing, Optional ByRef oRegGrp As InfoRepository.Registrations = Nothing)
    '    MyBase.New()

    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()
    '    Me.MdiParent = frmMusterContainer
    '    oReg = New InfoRepository.Registrations
    '    oReg = oRegGrp
    '    'Add any initialization after the InitializeComponent() call
    '    ' AddHandler TransferOwnership.Load, AddressOf mstContainer.loadTransferOwner
    'End Sub
    Public Sub New(ByVal ownID As Integer, Optional ByVal owner As MUSTER.BusinessLogic.pOwner = Nothing)
        MyBase.New()
        bolLoading = True
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        nOwnerID = ownID
        If owner Is Nothing Then
            pOwn = New MUSTER.BusinessLogic.pOwner
        Else
            pOwn = owner
        End If
        pOwn.Retrieve(nOwnerID)
        pOwn2 = New MUSTER.BusinessLogic.pOwner
        If pOwn.PersonID = 0 Then
            Me.txtBoxOwnerLeft.Text = IIf(IsNothing(pOwn.BPersona.Company), String.Empty, CStr(pOwn.BPersona.Company))
        Else
            Me.txtBoxOwnerLeft.Text = IIf(IsNothing(pOwn.Persona.Prefix), "", pOwn.Persona.Prefix + " ") + pOwn.Persona.FirstName + IIf(IsNothing(pOwn.Persona.MiddleName), "", pOwn.Persona.MiddleName + " ") + pOwn.Persona.LastName + IIf(IsNothing(pOwn.Persona.Suffix), "", " " + pOwn.Persona.Suffix)
        End If
        Me.txtBoxOwnerLeft.ReadOnly = True
        bolLoading = False
    End Sub
    'Public Sub New(ByRef frm As Registration, Optional ByRef frmMusterContainer As MusterContainer = Nothing, Optional ByRef oRegGrp As MUSTER.BusinessLogic.pRegistration = Nothing, Optional ByRef pOwner As MUSTER.BusinessLogic.pOwner = Nothing)
    '    MyBase.New()

    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()
    '    frmRegistration = frm
    '    Me.MdiParent = frmMusterContainer
    '    pOwn = pOwner
    '    pOwn2 = New MUSTER.BusinessLogic.pOwner
    '    oRegGroupRight = New MUSTER.BusinessLogic.pRegistration
    '    oRegGroupRight = oRegGrp
    '    If pOwn.PersonID = 0 Then
    '        Me.txtBoxOwnerLeft.Text = IIf(IsNothing(pOwn.BPersona.Company), String.Empty, CStr(pOwn.BPersona.Company))
    '    Else
    '        Me.txtBoxOwnerLeft.Text = IIf(IsNothing(pOwn.Persona.Prefix), "", pOwn.Persona.Prefix + " ") + pOwn.Persona.FirstName + IIf(IsNothing(pOwn.Persona.MiddleName), "", pOwn.Persona.MiddleName + " ") + pOwn.Persona.LastName + IIf(IsNothing(pOwn.Persona.Suffix), "", " " + pOwn.Persona.Suffix)
    '    End If
    '    Me.txtBoxOwnerLeft.ReadOnly = True
    'End Sub
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
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents grpBxFacDetail As System.Windows.Forms.GroupBox
    Friend WithEvents lstViewFacilityDetail As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnConfirmTransfer As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents chkBxNewOwnerSignature As System.Windows.Forms.CheckBox
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlMainLeft As System.Windows.Forms.Panel
    Friend WithEvents pnlRight As System.Windows.Forms.Panel
    Friend WithEvents pnlLeftSub As System.Windows.Forms.Panel
    Friend WithEvents pnlMiddle As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lstViewFacLeft As System.Windows.Forms.ListView
    Friend WithEvents ColHeaderSelectFac1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColHeaderFac1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColHeaderFacAddr1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents pnlOwner1Top As System.Windows.Forms.Panel
    Friend WithEvents lblCapCompliant1 As System.Windows.Forms.Label
    Friend WithEvents btnOwner1 As System.Windows.Forms.Button
    Friend WithEvents txtBoxOwnerLeft As System.Windows.Forms.TextBox
    Friend WithEvents lblOwner1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lstViewFacRight As System.Windows.Forms.ListView
    Friend WithEvents ColHeaderSelectFac2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColHeaderFac2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColHeadFacAddr2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents pnlOwner2Top As System.Windows.Forms.Panel
    Friend WithEvents lblCapParticipantOwner2 As System.Windows.Forms.Label
    Friend WithEvents lblOwner2 As System.Windows.Forms.Label
    Friend WithEvents btnShiftRight As System.Windows.Forms.Button
    Friend WithEvents btnShiftLeft As System.Windows.Forms.Button
    Friend WithEvents pnlFacilitiesCount As System.Windows.Forms.Panel
    Friend WithEvents lblNoofFacilitiesForOwner As System.Windows.Forms.Label
    Friend WithEvents lblNoOfFacilitiesOwnerValue As System.Windows.Forms.Label
    Friend WithEvents lblNoOfFacilitiesOwnerPotential As System.Windows.Forms.Label
    Friend WithEvents lblNoOfFacilitiesOwnerPotentialValue As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblCAPParticipantValueOwner1 As System.Windows.Forms.Label
    Friend WithEvents lblCAPParticipantValueOwner2 As System.Windows.Forms.Label
    Friend WithEvents cmbOwnerRight As System.Windows.Forms.ComboBox
    Friend WithEvents Facility_Id As System.Windows.Forms.ColumnHeader
    Friend WithEvents CAP_Participant As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents City As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.grpBxFacDetail = New System.Windows.Forms.GroupBox
        Me.lstViewFacilityDetail = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader8 = New System.Windows.Forms.ColumnHeader
        Me.btnConfirmTransfer = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.chkBxNewOwnerSignature = New System.Windows.Forms.CheckBox
        Me.pnlFacilitiesCount = New System.Windows.Forms.Panel
        Me.lblNoOfFacilitiesOwnerPotentialValue = New System.Windows.Forms.Label
        Me.lblNoOfFacilitiesOwnerPotential = New System.Windows.Forms.Label
        Me.lblNoOfFacilitiesOwnerValue = New System.Windows.Forms.Label
        Me.lblNoofFacilitiesForOwner = New System.Windows.Forms.Label
        Me.pnlMainLeft = New System.Windows.Forms.Panel
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.pnlMiddle = New System.Windows.Forms.Panel
        Me.btnShiftLeft = New System.Windows.Forms.Button
        Me.btnShiftRight = New System.Windows.Forms.Button
        Me.pnlLeftSub = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lstViewFacLeft = New System.Windows.Forms.ListView
        Me.ColHeaderSelectFac1 = New System.Windows.Forms.ColumnHeader
        Me.Facility_Id = New System.Windows.Forms.ColumnHeader
        Me.ColHeaderFac1 = New System.Windows.Forms.ColumnHeader
        Me.ColHeaderFacAddr1 = New System.Windows.Forms.ColumnHeader
        Me.City = New System.Windows.Forms.ColumnHeader
        Me.CAP_Participant = New System.Windows.Forms.ColumnHeader
        Me.pnlOwner1Top = New System.Windows.Forms.Panel
        Me.lblCAPParticipantValueOwner1 = New System.Windows.Forms.Label
        Me.lblCapCompliant1 = New System.Windows.Forms.Label
        Me.btnOwner1 = New System.Windows.Forms.Button
        Me.txtBoxOwnerLeft = New System.Windows.Forms.TextBox
        Me.lblOwner1 = New System.Windows.Forms.Label
        Me.pnlRight = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lstViewFacRight = New System.Windows.Forms.ListView
        Me.ColHeaderSelectFac2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader7 = New System.Windows.Forms.ColumnHeader
        Me.ColHeaderFac2 = New System.Windows.Forms.ColumnHeader
        Me.ColHeadFacAddr2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader9 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.pnlOwner2Top = New System.Windows.Forms.Panel
        Me.cmbOwnerRight = New System.Windows.Forms.ComboBox
        Me.lblCAPParticipantValueOwner2 = New System.Windows.Forms.Label
        Me.lblCapParticipantOwner2 = New System.Windows.Forms.Label
        Me.lblOwner2 = New System.Windows.Forms.Label
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.pnlBottom.SuspendLayout()
        Me.grpBxFacDetail.SuspendLayout()
        Me.pnlFacilitiesCount.SuspendLayout()
        Me.pnlMainLeft.SuspendLayout()
        Me.pnlMiddle.SuspendLayout()
        Me.pnlLeftSub.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.pnlOwner1Top.SuspendLayout()
        Me.pnlRight.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.pnlOwner2Top.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.grpBxFacDetail)
        Me.pnlBottom.Controls.Add(Me.btnConfirmTransfer)
        Me.pnlBottom.Controls.Add(Me.btnCancel)
        Me.pnlBottom.Controls.Add(Me.chkBxNewOwnerSignature)
        Me.pnlBottom.Controls.Add(Me.pnlFacilitiesCount)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.DockPadding.Left = 10
        Me.pnlBottom.Location = New System.Drawing.Point(0, 438)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(800, 160)
        Me.pnlBottom.TabIndex = 0
        '
        'grpBxFacDetail
        '
        Me.grpBxFacDetail.Controls.Add(Me.lstViewFacilityDetail)
        Me.grpBxFacDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpBxFacDetail.Location = New System.Drawing.Point(10, 18)
        Me.grpBxFacDetail.Name = "grpBxFacDetail"
        Me.grpBxFacDetail.Size = New System.Drawing.Size(790, 78)
        Me.grpBxFacDetail.TabIndex = 15
        Me.grpBxFacDetail.TabStop = False
        Me.grpBxFacDetail.Text = "Facility Detail"
        Me.grpBxFacDetail.Visible = False
        '
        'lstViewFacilityDetail
        '
        Me.lstViewFacilityDetail.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader8})
        Me.lstViewFacilityDetail.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstViewFacilityDetail.FullRowSelect = True
        Me.lstViewFacilityDetail.GridLines = True
        Me.lstViewFacilityDetail.Location = New System.Drawing.Point(3, 16)
        Me.lstViewFacilityDetail.Name = "lstViewFacilityDetail"
        Me.lstViewFacilityDetail.Size = New System.Drawing.Size(784, 59)
        Me.lstViewFacilityDetail.TabIndex = 4
        Me.lstViewFacilityDetail.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Owner"
        Me.ColumnHeader1.Width = 94
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Facility Name"
        Me.ColumnHeader2.Width = 137
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Address"
        Me.ColumnHeader3.Width = 231
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Lust Site"
        Me.ColumnHeader4.Width = 112
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Financial Site"
        Me.ColumnHeader5.Width = 206
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "index"
        Me.ColumnHeader8.Width = 0
        '
        'btnConfirmTransfer
        '
        Me.btnConfirmTransfer.Location = New System.Drawing.Point(416, 128)
        Me.btnConfirmTransfer.Name = "btnConfirmTransfer"
        Me.btnConfirmTransfer.Size = New System.Drawing.Size(104, 23)
        Me.btnConfirmTransfer.TabIndex = 14
        Me.btnConfirmTransfer.Text = "Save Transfer"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(528, 128)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Close"
        '
        'chkBxNewOwnerSignature
        '
        Me.chkBxNewOwnerSignature.Location = New System.Drawing.Point(14, 122)
        Me.chkBxNewOwnerSignature.Name = "chkBxNewOwnerSignature"
        Me.chkBxNewOwnerSignature.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkBxNewOwnerSignature.Size = New System.Drawing.Size(200, 24)
        Me.chkBxNewOwnerSignature.TabIndex = 12
        Me.chkBxNewOwnerSignature.Text = "New Owner's Signature Available"
        Me.chkBxNewOwnerSignature.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlFacilitiesCount
        '
        Me.pnlFacilitiesCount.Controls.Add(Me.lblNoOfFacilitiesOwnerPotentialValue)
        Me.pnlFacilitiesCount.Controls.Add(Me.lblNoOfFacilitiesOwnerPotential)
        Me.pnlFacilitiesCount.Controls.Add(Me.lblNoOfFacilitiesOwnerValue)
        Me.pnlFacilitiesCount.Controls.Add(Me.lblNoofFacilitiesForOwner)
        Me.pnlFacilitiesCount.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFacilitiesCount.Location = New System.Drawing.Point(10, 0)
        Me.pnlFacilitiesCount.Name = "pnlFacilitiesCount"
        Me.pnlFacilitiesCount.Size = New System.Drawing.Size(790, 18)
        Me.pnlFacilitiesCount.TabIndex = 16
        '
        'lblNoOfFacilitiesOwnerPotentialValue
        '
        Me.lblNoOfFacilitiesOwnerPotentialValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilitiesOwnerPotentialValue.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblNoOfFacilitiesOwnerPotentialValue.Location = New System.Drawing.Point(750, 0)
        Me.lblNoOfFacilitiesOwnerPotentialValue.Name = "lblNoOfFacilitiesOwnerPotentialValue"
        Me.lblNoOfFacilitiesOwnerPotentialValue.Size = New System.Drawing.Size(40, 18)
        Me.lblNoOfFacilitiesOwnerPotentialValue.TabIndex = 3
        '
        'lblNoOfFacilitiesOwnerPotential
        '
        Me.lblNoOfFacilitiesOwnerPotential.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilitiesOwnerPotential.Location = New System.Drawing.Point(629, 0)
        Me.lblNoOfFacilitiesOwnerPotential.Name = "lblNoOfFacilitiesOwnerPotential"
        Me.lblNoOfFacilitiesOwnerPotential.Size = New System.Drawing.Size(100, 18)
        Me.lblNoOfFacilitiesOwnerPotential.TabIndex = 2
        Me.lblNoOfFacilitiesOwnerPotential.Text = "No. of Facilities:"
        '
        'lblNoOfFacilitiesOwnerValue
        '
        Me.lblNoOfFacilitiesOwnerValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilitiesOwnerValue.Location = New System.Drawing.Point(120, 0)
        Me.lblNoOfFacilitiesOwnerValue.Name = "lblNoOfFacilitiesOwnerValue"
        Me.lblNoOfFacilitiesOwnerValue.Size = New System.Drawing.Size(40, 18)
        Me.lblNoOfFacilitiesOwnerValue.TabIndex = 1
        '
        'lblNoofFacilitiesForOwner
        '
        Me.lblNoofFacilitiesForOwner.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoofFacilitiesForOwner.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoofFacilitiesForOwner.Location = New System.Drawing.Point(0, 0)
        Me.lblNoofFacilitiesForOwner.Name = "lblNoofFacilitiesForOwner"
        Me.lblNoofFacilitiesForOwner.Size = New System.Drawing.Size(100, 18)
        Me.lblNoofFacilitiesForOwner.TabIndex = 0
        Me.lblNoofFacilitiesForOwner.Text = "No. of Facilities:"
        '
        'pnlMainLeft
        '
        Me.pnlMainLeft.Controls.Add(Me.Splitter2)
        Me.pnlMainLeft.Controls.Add(Me.pnlMiddle)
        Me.pnlMainLeft.Controls.Add(Me.pnlLeftSub)
        Me.pnlMainLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlMainLeft.DockPadding.Left = 10
        Me.pnlMainLeft.Location = New System.Drawing.Point(0, 0)
        Me.pnlMainLeft.Name = "pnlMainLeft"
        Me.pnlMainLeft.Size = New System.Drawing.Size(440, 438)
        Me.pnlMainLeft.TabIndex = 1
        '
        'Splitter2
        '
        Me.Splitter2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Splitter2.Location = New System.Drawing.Point(386, 0)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(3, 438)
        Me.Splitter2.TabIndex = 2
        Me.Splitter2.TabStop = False
        '
        'pnlMiddle
        '
        Me.pnlMiddle.Controls.Add(Me.btnShiftLeft)
        Me.pnlMiddle.Controls.Add(Me.btnShiftRight)
        Me.pnlMiddle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMiddle.Location = New System.Drawing.Point(386, 0)
        Me.pnlMiddle.Name = "pnlMiddle"
        Me.pnlMiddle.Size = New System.Drawing.Size(54, 438)
        Me.pnlMiddle.TabIndex = 1
        '
        'btnShiftLeft
        '
        Me.btnShiftLeft.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnShiftLeft.Location = New System.Drawing.Point(11, 167)
        Me.btnShiftLeft.Name = "btnShiftLeft"
        Me.btnShiftLeft.Size = New System.Drawing.Size(32, 32)
        Me.btnShiftLeft.TabIndex = 1
        Me.btnShiftLeft.Text = "<<"
        '
        'btnShiftRight
        '
        Me.btnShiftRight.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnShiftRight.Location = New System.Drawing.Point(10, 128)
        Me.btnShiftRight.Name = "btnShiftRight"
        Me.btnShiftRight.Size = New System.Drawing.Size(32, 32)
        Me.btnShiftRight.TabIndex = 0
        Me.btnShiftRight.Text = ">>"
        '
        'pnlLeftSub
        '
        Me.pnlLeftSub.Controls.Add(Me.GroupBox1)
        Me.pnlLeftSub.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlLeftSub.Location = New System.Drawing.Point(10, 0)
        Me.pnlLeftSub.Name = "pnlLeftSub"
        Me.pnlLeftSub.Size = New System.Drawing.Size(376, 438)
        Me.pnlLeftSub.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lstViewFacLeft)
        Me.GroupBox1.Controls.Add(Me.pnlOwner1Top)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(376, 438)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'lstViewFacLeft
        '
        Me.lstViewFacLeft.AllowColumnReorder = True
        Me.lstViewFacLeft.CheckBoxes = True
        Me.lstViewFacLeft.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColHeaderSelectFac1, Me.Facility_Id, Me.ColHeaderFac1, Me.ColHeaderFacAddr1, Me.City, Me.CAP_Participant})
        Me.lstViewFacLeft.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstViewFacLeft.FullRowSelect = True
        Me.lstViewFacLeft.GridLines = True
        Me.lstViewFacLeft.Location = New System.Drawing.Point(3, 88)
        Me.lstViewFacLeft.Name = "lstViewFacLeft"
        Me.lstViewFacLeft.Size = New System.Drawing.Size(370, 347)
        Me.lstViewFacLeft.TabIndex = 13
        Me.lstViewFacLeft.View = System.Windows.Forms.View.Details
        '
        'ColHeaderSelectFac1
        '
        Me.ColHeaderSelectFac1.Text = ""
        Me.ColHeaderSelectFac1.Width = 22
        '
        'Facility_Id
        '
        Me.Facility_Id.Text = "Facility ID"
        Me.Facility_Id.Width = 76
        '
        'ColHeaderFac1
        '
        Me.ColHeaderFac1.Text = "Facility Name"
        Me.ColHeaderFac1.Width = 94
        '
        'ColHeaderFacAddr1
        '
        Me.ColHeaderFacAddr1.Text = "Address"
        Me.ColHeaderFacAddr1.Width = 77
        '
        'City
        '
        Me.City.Text = "City"
        Me.City.Width = 75
        '
        'CAP_Participant
        '
        Me.CAP_Participant.Text = "CAP Participant"
        Me.CAP_Participant.Width = 0
        '
        'pnlOwner1Top
        '
        Me.pnlOwner1Top.Controls.Add(Me.lblCAPParticipantValueOwner1)
        Me.pnlOwner1Top.Controls.Add(Me.lblCapCompliant1)
        Me.pnlOwner1Top.Controls.Add(Me.btnOwner1)
        Me.pnlOwner1Top.Controls.Add(Me.txtBoxOwnerLeft)
        Me.pnlOwner1Top.Controls.Add(Me.lblOwner1)
        Me.pnlOwner1Top.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwner1Top.Location = New System.Drawing.Point(3, 16)
        Me.pnlOwner1Top.Name = "pnlOwner1Top"
        Me.pnlOwner1Top.Size = New System.Drawing.Size(370, 72)
        Me.pnlOwner1Top.TabIndex = 12
        '
        'lblCAPParticipantValueOwner1
        '
        Me.lblCAPParticipantValueOwner1.Location = New System.Drawing.Point(104, 50)
        Me.lblCAPParticipantValueOwner1.Name = "lblCAPParticipantValueOwner1"
        Me.lblCAPParticipantValueOwner1.TabIndex = 4
        '
        'lblCapCompliant1
        '
        Me.lblCapCompliant1.Location = New System.Drawing.Point(8, 50)
        Me.lblCapCompliant1.Name = "lblCapCompliant1"
        Me.lblCapCompliant1.Size = New System.Drawing.Size(88, 23)
        Me.lblCapCompliant1.TabIndex = 3
        Me.lblCapCompliant1.Text = "CAP Participant:"
        '
        'btnOwner1
        '
        Me.btnOwner1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwner1.Location = New System.Drawing.Point(269, 24)
        Me.btnOwner1.Name = "btnOwner1"
        Me.btnOwner1.Size = New System.Drawing.Size(24, 23)
        Me.btnOwner1.TabIndex = 1
        Me.btnOwner1.Text = "..."
        Me.btnOwner1.Visible = False
        '
        'txtBoxOwnerLeft
        '
        Me.txtBoxOwnerLeft.Location = New System.Drawing.Point(8, 24)
        Me.txtBoxOwnerLeft.Name = "txtBoxOwnerLeft"
        Me.txtBoxOwnerLeft.Size = New System.Drawing.Size(256, 20)
        Me.txtBoxOwnerLeft.TabIndex = 0
        Me.txtBoxOwnerLeft.Text = ""
        '
        'lblOwner1
        '
        Me.lblOwner1.Location = New System.Drawing.Point(7, 7)
        Me.lblOwner1.Name = "lblOwner1"
        Me.lblOwner1.TabIndex = 2
        Me.lblOwner1.Text = "Current Owner"
        '
        'pnlRight
        '
        Me.pnlRight.Controls.Add(Me.GroupBox2)
        Me.pnlRight.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlRight.DockPadding.Left = 4
        Me.pnlRight.Location = New System.Drawing.Point(440, 0)
        Me.pnlRight.Name = "pnlRight"
        Me.pnlRight.Size = New System.Drawing.Size(360, 438)
        Me.pnlRight.TabIndex = 2
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lstViewFacRight)
        Me.GroupBox2.Controls.Add(Me.pnlOwner2Top)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(4, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(356, 438)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'lstViewFacRight
        '
        Me.lstViewFacRight.AllowColumnReorder = True
        Me.lstViewFacRight.CheckBoxes = True
        Me.lstViewFacRight.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColHeaderSelectFac2, Me.ColumnHeader7, Me.ColHeaderFac2, Me.ColHeadFacAddr2, Me.ColumnHeader9, Me.ColumnHeader6})
        Me.lstViewFacRight.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstViewFacRight.FullRowSelect = True
        Me.lstViewFacRight.GridLines = True
        Me.lstViewFacRight.Location = New System.Drawing.Point(3, 88)
        Me.lstViewFacRight.Name = "lstViewFacRight"
        Me.lstViewFacRight.Size = New System.Drawing.Size(350, 347)
        Me.lstViewFacRight.TabIndex = 10
        Me.lstViewFacRight.View = System.Windows.Forms.View.Details
        '
        'ColHeaderSelectFac2
        '
        Me.ColHeaderSelectFac2.Text = ""
        Me.ColHeaderSelectFac2.Width = 22
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Facility ID"
        Me.ColumnHeader7.Width = 71
        '
        'ColHeaderFac2
        '
        Me.ColHeaderFac2.Text = "Facility Name"
        Me.ColHeaderFac2.Width = 100
        '
        'ColHeadFacAddr2
        '
        Me.ColHeadFacAddr2.Text = "Address"
        Me.ColHeadFacAddr2.Width = 175
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "City"
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "CAP Participant"
        Me.ColumnHeader6.Width = 0
        '
        'pnlOwner2Top
        '
        Me.pnlOwner2Top.Controls.Add(Me.cmbOwnerRight)
        Me.pnlOwner2Top.Controls.Add(Me.lblCAPParticipantValueOwner2)
        Me.pnlOwner2Top.Controls.Add(Me.lblCapParticipantOwner2)
        Me.pnlOwner2Top.Controls.Add(Me.lblOwner2)
        Me.pnlOwner2Top.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwner2Top.Location = New System.Drawing.Point(3, 16)
        Me.pnlOwner2Top.Name = "pnlOwner2Top"
        Me.pnlOwner2Top.Size = New System.Drawing.Size(350, 72)
        Me.pnlOwner2Top.TabIndex = 9
        '
        'cmbOwnerRight
        '
        Me.cmbOwnerRight.Location = New System.Drawing.Point(8, 24)
        Me.cmbOwnerRight.Name = "cmbOwnerRight"
        Me.cmbOwnerRight.Size = New System.Drawing.Size(296, 21)
        Me.cmbOwnerRight.TabIndex = 8
        '
        'lblCAPParticipantValueOwner2
        '
        Me.lblCAPParticipantValueOwner2.Location = New System.Drawing.Point(104, 50)
        Me.lblCAPParticipantValueOwner2.Name = "lblCAPParticipantValueOwner2"
        Me.lblCAPParticipantValueOwner2.TabIndex = 7
        '
        'lblCapParticipantOwner2
        '
        Me.lblCapParticipantOwner2.Location = New System.Drawing.Point(8, 50)
        Me.lblCapParticipantOwner2.Name = "lblCapParticipantOwner2"
        Me.lblCapParticipantOwner2.Size = New System.Drawing.Size(88, 23)
        Me.lblCapParticipantOwner2.TabIndex = 6
        Me.lblCapParticipantOwner2.Text = "CAP Participant:"
        '
        'lblOwner2
        '
        Me.lblOwner2.Location = New System.Drawing.Point(8, 7)
        Me.lblOwner2.Name = "lblOwner2"
        Me.lblOwner2.TabIndex = 5
        Me.lblOwner2.Text = "Potential Owner"
        '
        'Splitter1
        '
        Me.Splitter1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Splitter1.Location = New System.Drawing.Point(440, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 438)
        Me.Splitter1.TabIndex = 3
        Me.Splitter1.TabStop = False
        '
        'TransferOwnership
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(800, 598)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.pnlRight)
        Me.Controls.Add(Me.pnlMainLeft)
        Me.Controls.Add(Me.pnlBottom)
        Me.Name = "TransferOwnership"
        Me.Text = "Registration - Transfer Ownership"
        Me.pnlBottom.ResumeLayout(False)
        Me.grpBxFacDetail.ResumeLayout(False)
        Me.pnlFacilitiesCount.ResumeLayout(False)
        Me.pnlMainLeft.ResumeLayout(False)
        Me.pnlMiddle.ResumeLayout(False)
        Me.pnlLeftSub.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.pnlOwner1Top.ResumeLayout(False)
        Me.pnlRight.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.pnlOwner2Top.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnConfirmTransfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConfirmTransfer.Click
        Dim lstViewItem As ListViewItem
        Try
            If cmbOwnerRight.SelectedValue = 0 Then
                For Each lstViewItem In Me.lstViewFacRight.CheckedItems
                    lstViewItem.Checked = False
                    Me.lstViewFacRight.Items.RemoveAt(lstViewItem.Index)
                    Me.lstViewFacLeft.Items.Add(lstViewItem)
                Next
                Me.lblNoOfFacilitiesOwnerValue.Text = lstViewFacLeft.Items.Count
                Me.lblNoOfFacilitiesOwnerPotentialValue.Text = lstViewFacRight.Items.Count
                MsgBox("OwnerName is missing")
                Exit Sub
            End If

            If Not lstViewFacRight.CheckedItems.Count > 0 Then
                MsgBox("No Facilities to Transfer")
                Exit Sub
            End If

            If lstViewFacRight.Items.Count > 0 And lstViewFacRight.CheckedItems.Count > 0 Then
                ' check if fees can create credit memo (cannot create credit memo if there is no invoice number from bp2k)
                Dim strErr As String = String.Empty
                Dim ds As DataSet
                Dim bolhasBalance As Boolean = False
                For Each lstViewItem In Me.lstViewFacRight.CheckedItems
                    bolhasBalance = False
                    ds = pOwn2.RunSQLQuery("SELECT	isnull(Sum(cast(isnull(debit,0) as money)) - sum(cast(isnull(Credit,0) as money)),0) FROM vFEES_FacilityFeeByTransaction Where INV_DATE is not null and Facility_ID = " + lstViewItem.SubItems(1).Text)
                    If ds.Tables(0).Rows.Count > 0 Then
                        If ds.Tables(0).Rows(0)(0) > 0 Then
                            bolhasBalance = True
                        End If
                    End If
                    If bolhasBalance Then
                        ds = pOwn2.RunSQLQuery("select inv_number from vFees_DetailLine_Subcalc_BilledAmount where facility_id = " + lstViewItem.SubItems(1).Text)
                        If ds.Tables(0).Rows.Count <= 0 Then
                            strErr += lstViewItem.SubItems(1).Text + ","
                        End If
                    End If
                Next
                If strErr <> String.Empty Then
                    strErr = strErr.Trim.TrimEnd(",")
                    MsgBox("Cannot transfer the following fac(s)" + vbCrLf + strErr + vbCrLf + _
                        "Cannot create credit memo for invoices without Invoice Number", MsgBoxStyle.Exclamation, "Transfer interrupted")
                    Exit Sub
                End If

                oRegGroupRight.RetrieveByOwnerID(cmbOwnerRight.SelectedValue)
                If oRegGroupRight.ID <= 0 Then
                    If Not chkBxNewOwnerSignature.Checked Then
                        oRegGroupRight.OWNER_ID = cmbOwnerRight.SelectedValue
                        oRegGroupRight.DATE_STARTED = Now
                        oRegGroupRight.DATE_COMPLETED = CDate("01/01/0001")
                        oRegGroupRight.COMPLETED = False
                        oRegGroupRight.Deleted = False
                        oRegGroupRight.Save()
                    End If
                End If

                Dim bolOwnerLeftisOpen As Boolean = False
                Dim bolOwnerRightisOpen As Boolean = False
                Dim frmRegRight As Registration
                Dim frmRegLeft As Registration
                Dim facsCol As New MUSTER.Info.FacilityCollection

                bolOwnerRightisOpen = False
                pOwn2 = New MUSTER.BusinessLogic.pOwner
                For Each frmChild As Form In Me.MdiParent.MdiChildren
                    If frmChild.Name.ToUpper = "REGISTRATION" Then
                        If CType(frmChild, Registration).nOwnerID = cmbOwnerRight.SelectedValue Then
                            frmRegRight = CType(frmChild, Registration)
                            'pOwn2 = frmRegRight.pOwn
                            'frmReg.bolRefreshFacilitiesGrid = True
                            bolOwnerRightisOpen = True
                        ElseIf CType(frmChild, Registration).nOwnerID = nOwnerID Then
                            frmRegLeft = CType(frmChild, Registration)
                            bolOwnerLeftisOpen = True
                        End If
                    End If
                    If bolOwnerLeftisOpen And bolOwnerRightisOpen Then Exit For
                Next

                Dim adviceID As Integer = 0

                For Each lstViewItem In Me.lstViewFacRight.CheckedItems
                    pOwn2.Retrieve(cmbOwnerRight.SelectedValue)
                    pOwn2.Facilities.Retrieve(pOwn2.OwnerInfo, CInt(lstViewItem.SubItems(1).Text), , "FACILITY")
                    pOwn2.Facilities.OwnerID = cmbOwnerRight.SelectedValue

                    pOwn2.Facilities.DateTransferred = Now
                    If chkBxNewOwnerSignature.Checked Then
                        pOwn2.Facilities.SignatureOnNF = True
                    Else
                        pOwn2.Facilities.SignatureOnNF = False
                    End If
                    pOwn2.Facilities.ModifiedBy = MusterContainer.AppUser.ID
                    ' #2802
                    If pOwn2.CAPParticipationLevel.StartsWith("FULL") Or pOwn2.CAPParticipationLevel.StartsWith("PARTIAL") Then
                        pOwn2.Facilities.CAPCandidate = True
                    End If
                    pOwn2.Facilities.Save(CType(UIUtilsGen.ModuleID.Registration, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False, False, adviceID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    ' delete sig req
                    DeleteRegistrationActivity(pOwn.ID, pOwn2.Facilities.ID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.SignatureRequired)

                    ' move all activites related to the facility to the new owner
                    MoveTankRegistrationActivitiesToNewOwner(pOwn2.ID, pOwn.ID, pOwn2.Facilities.ID)

                    ' #2873 - put activity transfer owner to new owner
                    PutRegistrationActivity(pOwn2.ID, pOwn2.Facilities.ID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.TransferOwnership)

                    ' Check for tosi tanks
                    ' if exists add/enable tosi activity for the tanks
                    ds = pOwn2.RunSQLQuery("select tank_id from tblreg_tank where deleted = 0 and tankstatus = 429 and facility_id = " + pOwn2.Facilities.ID.ToString)
                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each dr As DataRow In ds.Tables(0).Rows
                            PutRegistrationActivity(pOwn2.ID, dr("tank_id"), UIUtilsGen.EntityTypes.Tank, UIUtilsGen.ActivityTypes.TankStatusTOSI)
                        Next
                    End If


                    If chkBxNewOwnerSignature.Checked Then
                        If bolOwnerRightisOpen Then
                            frmRegRight.CheckForRegistrationActivity()
                        End If
                    Else
                        If bolOwnerRightisOpen Then
                            frmRegRight.PutRegistrationActivity(pOwn2.Facilities.ID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.SignatureRequired)
                        Else
                            PutRegistrationActivity(pOwn2.ID, pOwn2.Facilities.ID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ActivityTypes.SignatureRequired)
                        End If
                    End If

                    If Not facsCol.Contains(pOwn2.Facilities.ID) Then
                        facsCol.Add(pOwn2.Facilities.Facility)
                    End If

                Next

                '   pOwn2.TransferOwnerBilling(facs.ToString(), nOwnerID, pOwn2.ID)



                If bolOwnerRightisOpen Then
                    Dim myGuid As System.Guid
                    frmRegRight.lblFacilityIDValue.Text = String.Empty
                    myGuid = frmRegRight.MyGuid
                    MusterContainer.AppSemaphores.Remove(myGuid.ToString)
                    MusterContainer.AppSemaphores.Retrieve(myGuid.ToString, "WindowName", "Registration", "Registration")
                    MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", myGuid, "Registration")
                    pOwn2.OwnerInfo.facilityCollection = New MUSTER.Info.FacilityCollection

                    If frmRegRight.tbCntrlRegistration.SelectedTab.Name = frmRegRight.tbPageOwnerDetail.Name Then
                        frmRegRight.PopulateOwner(pOwn2.ID, True)
                    Else
                        frmRegRight.tbCntrlRegistration.SelectedTab = frmRegRight.tbPageOwnerDetail
                    End If
                End If

                ' Generate Transfer Letter for Old Owner
                Dim regLetters As New Reg_Letters
                regLetters.GenerateTransferAcknowledgementLetter(nOwnerID, facsCol, pOwn, False)

                '' Generate Transfer Letter for New Owner
                'regLetters.GenerateTransferAcknowledgementLetter(cmbOwnerRight.SelectedValue, facsCol, pOwn2, True)

                ' after the facility is transferred, it no longer belongs to the current owner
                pOwn.OwnerInfo.facilityCollection = New MUSTER.Info.FacilityCollection

                If bolOwnerLeftisOpen Then
                    Dim myGuid As System.Guid
                    frmRegLeft.lblFacilityIDValue.Text = String.Empty
                    myGuid = frmRegLeft.MyGuid
                    MusterContainer.AppSemaphores.Remove(myGuid.ToString)
                    MusterContainer.AppSemaphores.Retrieve(myGuid.ToString, "WindowName", "Registration", "Registration")
                    MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", myGuid, "Registration")
                    frmRegLeft.pOwn.OwnerInfo.facilityCollection = New MUSTER.Info.FacilityCollection
                    frmRegLeft.nFacilityID = 0
                    If frmRegLeft.tbCntrlRegistration.SelectedTab.Name = frmRegLeft.tbPageOwnerDetail.Name Then
                        frmRegLeft.PopulateOwner(nOwnerID, True)
                    Else
                        frmRegLeft.tbCntrlRegistration.SelectedTab = frmRegLeft.tbPageOwnerDetail
                    End If
                End If

                Me.Close()

            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception(" Transfer Ownership Failed: " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub DeleteRegistrationActivity(ByVal oldOwnerID As Integer, ByVal EntityID As Integer, ByVal EntityType As Integer, ByVal Activity As String)
        Try
            Dim bolDeleteActivity As Boolean = False

            Dim oRegGroupLeft As New MUSTER.BusinessLogic.pRegistration
            oRegGroupLeft.RetrieveByOwnerID(oldOwnerID)

            If oRegGroupLeft.ID > 0 Then
                oRegGroupLeft.Activity.Col = oRegGroupLeft.Activity.RetrieveByRegID(oRegGroupLeft.ID)
                If oRegGroupLeft.Activity.Col.Count > 0 Then
                    For Each oRegActinfo In oRegGroupLeft.Activity.Col.Values
                        If oRegActinfo.EntityId = EntityID And _
                            oRegActinfo.EntityType = EntityType And _
                            oRegActinfo.ActivityDesc = Activity Then
                            oRegGroupLeft.Activity.RegActivityInfo = oRegActinfo
                            bolDeleteActivity = True
                            Exit For
                        End If
                    Next
                End If
            End If



            If bolDeleteActivity Then
                ' delete cal entry - taken care in trigger
                oRegGroupLeft.Activity.Processed = True
                oRegGroupLeft.Activity.Deleted = True
                oRegGroupLeft.Activity.Save()

                oRegGroupLeft.Activity.Col.Remove(oRegGroupLeft.Activity.RegActionIndex)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub PutRegistrationActivity(ByVal ownID As Integer, ByVal EntityID As Integer, ByVal EntityType As Integer, ByVal Activity As String)
        Try
            Dim bolAddActivity As Boolean = True
            Dim bolShowMsg As Boolean = True

            If oRegGroupRight.OWNER_ID <> ownID Then
                oRegGroupRight = New MUSTER.BusinessLogic.pRegistration
                oRegGroupRight.RetrieveByOwnerID(ownID)
            End If

            If oRegGroupRight.ID <= 0 Then
                oRegGroupRight.OWNER_ID = ownID
                oRegGroupRight.DATE_STARTED = Now
                oRegGroupRight.DATE_COMPLETED = CDate("01/01/0001")
                oRegGroupRight.COMPLETED = False
                oRegGroupRight.Deleted = False
                oRegGroupRight.Save()
            End If

            oRegGroupRight.Activity.Col = oRegGroupRight.Activity.RetrieveByRegID(oRegGroupRight.ID)

            If oRegGroupRight.Activity.Col.Count > 0 Then
                bolShowMsg = False
                For Each oRegActinfo In oRegGroupRight.Activity.Col.Values
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
                        calDescText = ownID.ToString + " Incomplete Registration - New Owner Added"
                        'Case UIUtilsGen.ActivityTypes.TransferAcknowledgement
                        '    calDescText = ownID.ToString + " Incomplete Registration - Facility Transfer Acknowledgement"
                    Case UIUtilsGen.ActivityTypes.TransferOwnership
                        calDescText = ownID.ToString + " Incomplete Registration - Facility Transferred"
                    Case UIUtilsGen.ActivityTypes.UpComingInstall
                        calDescText = ownID.ToString + " Incomplete Registration - Upcoming Install"
                    Case UIUtilsGen.ActivityTypes.SignatureRequired
                        calDescText = ownID.ToString + " Incomplete Registration - Signature Required"
                    Case UIUtilsGen.ActivityTypes.TankStatusTOSI
                        calDescText = ownID.ToString + " Incomplete Registration - Tank Status changed to TOSI"
                    Case UIUtilsGen.ActivityTypes.AddTank
                        calDescText = ownID.ToString + " Incomplete Registration - New Tank Added"
                End Select

                Dim mc As MusterContainer = Me.MdiParent
                If mc Is Nothing Then mc = New MusterContainer
                mc.pCalendar.Add(New MUSTER.Info.CalendarInfo(0, Now, DateAdd(DateInterval.Day, 30, Now), 0, calDescText, mc.AppUser.ID, "SYSTEM", "", False, True, False, False, mc.AppUser.ID, Now, String.Empty, CDate("01/01/0001"), EntityType, EntityID))
                mc.pCalendar.Save()
                'Adding Registration Activity Detail
                oRegActinfo = New MUSTER.Info.RegistrationActivityInfo(0, _
                                                            oRegGroupRight.ID, _
                                                            EntityType, _
                                                            EntityID, _
                                                            MusterContainer.AppUser.ID, _
                                                            Activity, _
                                                            False, _
                                                            Now(), _
                                                            mc.pCalendar.CalendarId)
                oRegGroupRight.Activity.Add(oRegActinfo)
                oRegGroupRight.Activity.Save()
            End If

            If bolShowMsg Then
                MsgBox("Placing owner (" + ownID.ToString + ") in registration mode.", MsgBoxStyle.Information & MsgBoxStyle.OKOnly, "Registration Initiated")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub MoveTankRegistrationActivitiesToNewOwner(ByVal newOwnID As Integer, ByVal oldOwnID As Integer, ByVal facID As Integer)
        Try
            If oRegGroupRight.OWNER_ID <> newOwnID Then
                oRegGroupRight = New MUSTER.BusinessLogic.pRegistration
                oRegGroupRight.RetrieveByOwnerID(newOwnID)
            End If

            If oRegGroupRight.ID <= 0 Then
                ' if fee has inserted a row, need to capture that reg id
                oRegGroupRight.RetrieveByOwnerID(newOwnID)
                If oRegGroupRight.ID <= 0 Then
                    oRegGroupRight.OWNER_ID = newOwnID
                    oRegGroupRight.DATE_STARTED = Now
                    oRegGroupRight.DATE_COMPLETED = CDate("01/01/0001")
                    oRegGroupRight.COMPLETED = False
                    oRegGroupRight.Deleted = False
                    oRegGroupRight.Save()
                End If
            End If

            Dim oRegGroupLeft As New MUSTER.BusinessLogic.pRegistration
            Dim alFacTanks As New ArrayList

            oRegGroupLeft.RetrieveByOwnerID(oldOwnID)
            If oRegGroupLeft.ID > 0 Then
                oRegGroupLeft.Activity.Col = oRegGroupLeft.Activity.RetrieveByRegID(oRegGroupLeft.ID)
                If oRegGroupLeft.Activity.Col.Count > 0 Then
                    ' get tanks for a given fac
                    Dim ds As DataSet = pOwn.RunSQLQuery("SELECT DISTINCT TANK_ID FROM TBLREG_TANK WHERE FACILITY_ID = " + facID.ToString)
                    If ds.Tables.Count > 0 Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            For Each dr As DataRow In ds.Tables(0).Rows
                                alFacTanks.Add(dr.Item("TANK_ID"))
                            Next
                        End If
                    End If

                    For Each oRegActinfo In oRegGroupLeft.Activity.Col.Values
                        If oRegActinfo.EntityType = UIUtilsGen.EntityTypes.Tank Or _
                            oRegActinfo.EntityType = UIUtilsGen.EntityTypes.Facility Then
                            If (alFacTanks.Contains(oRegActinfo.EntityId) And oRegActinfo.EntityType = UIUtilsGen.EntityTypes.Tank) Or _
                                (oRegActinfo.EntityId = facID And oRegActinfo.EntityType = UIUtilsGen.EntityTypes.Facility) Then
                                oRegActinfo.RegistrationID = oRegGroupRight.ID
                                If oRegActinfo.CalendarID > 0 Then
                                    Dim mc As MusterContainer = Me.MdiParent
                                    If mc Is Nothing Then mc = New MusterContainer
                                    mc.pCalendar.Retrieve(oRegActinfo.CalendarID)
                                    If mc.pCalendar.CalendarId > 0 Then
                                        Dim str As String = mc.pCalendar.TaskDescription
                                        If str.IndexOf("Incomplete") > -1 Then
                                            str = str.Substring(str.IndexOf("Incomplete"))
                                            str = newOwnID.ToString + " " + str
                                            mc.pCalendar.TaskDescription = str
                                            mc.pCalendar.Save()
                                        End If
                                    Else
                                        mc.pCalendar.Remove(mc.pCalendar.CalendarId)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    oRegGroupLeft.Save()
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub btnShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShiftRight.Click
        Dim lstViewItem As ListViewItem
        Try

            If Me.cmbOwnerRight.SelectedIndex = -1 Or cmbOwnerRight.SelectedIndex = 0 Then
                MsgBox("Select the Potential Owner to Transfer.")
                Me.cmbOwnerRight.Focus()
                Exit Sub
            End If

            If Not lstViewFacLeft.CheckedItems Is Nothing Then
                If lstViewFacLeft.CheckedItems.Count <= 0 Then
                    MsgBox("No Facilities selected to Transfer")
                    Exit Sub
                End If
            End If

            For Each lstViewItem In Me.lstViewFacLeft.CheckedItems
                'lstViewItem.Checked = True
                Me.lstViewFacLeft.Items.RemoveAt(lstViewItem.Index)
                Me.lstViewFacRight.Items.Add(lstViewItem)
            Next
            Me.lblNoOfFacilitiesOwnerValue.Text = lstViewFacLeft.Items.Count
            Me.lblNoOfFacilitiesOwnerPotentialValue.Text = lstViewFacRight.Items.Count

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Transfer Facilities: " + ex.Message, ex))
            MyErr.ShowDialog()
            'MsgBox("Cannot Transfer Facilities: " + ex.Message)
        End Try
    End Sub
    Private Sub btnShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShiftLeft.Click
        Dim lstViewItem As ListViewItem
        Dim drGlobal As DataRow
        Try
            'For Each lstViewItem In Me.lstViewFacRight.CheckedItems

            '    For Each drGlobal In dtGlobal.Rows
            '        If Not lstViewItem.SubItems(1).Text = drGlobal("FACILITY_ID") Then
            '            MsgBox("Cannot Transfer Facilities from Right hand side")
            '            Exit Sub
            '        End If
            '    Next

            '    lstViewItem.Checked = False
            '    Me.lstViewFacRight.Items.RemoveAt(lstViewItem.Index)
            '    Me.lstViewFacLeft.Items.Add(lstViewItem)
            'Next

            Dim nItemCount As Integer = 0
            For Each lstViewItem In Me.lstViewFacRight.CheckedItems
                For Each drGlobal In dtGlobal.Rows
                    If lstViewItem.SubItems(1).Text = drGlobal("FACILITY_ID") Then
                        nItemCount += 1
                        Exit For
                    End If
                Next
            Next

            If Me.lstViewFacRight.CheckedItems.Count <> nItemCount Then
                MsgBox("Cannot Transfer Facilities from Right hand side")
                Exit Sub
            End If

            For Each lstViewItem In Me.lstViewFacRight.CheckedItems
                lstViewItem.Checked = False
                Me.lstViewFacRight.Items.RemoveAt(lstViewItem.Index)
                Me.lstViewFacLeft.Items.Add(lstViewItem)
            Next

            Me.lblNoOfFacilitiesOwnerValue.Text = lstViewFacLeft.Items.Count
            Me.lblNoOfFacilitiesOwnerPotentialValue.Text = lstViewFacRight.Items.Count

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Transfer Facilities: " + ex.Message, ex))
            MyErr.ShowDialog()
            'MsgBox("Cannot Transfer Facilities: " + ex.Message)
        End Try
    End Sub
    Private Function PopulateFacilities(ByVal OwnerID As Int64) As DataTable

        Dim xColFacilities As MUSTER.Info.FacilityCollection
        Dim xFacInfo As MUSTER.Info.FacilityInfo
        Dim xAddressInfo As MUSTER.Info.AddressInfo
        Dim dtFacility As New DataTable
        Dim dr As DataRow
        Dim sList As New SortedList
        Dim i As Integer = 0
        Try
            pOwn2 = New MUSTER.BusinessLogic.pOwner
            xColFacilities = pOwn2.Facilities.GetAllInfo(OwnerID)
            pOwn2 = New MUSTER.BusinessLogic.pOwner
            dtFacility.Columns.Add("OWNER_ID")
            dtFacility.Columns.Add("FACILITY_ID")
            dtFacility.Columns.Add("CAP_PARTICIPANT")
            dtFacility.Columns.Add("FACILITYNAME")
            dtFacility.Columns.Add("ADDRESS")
            dtFacility.Columns.Add("CITY")
            For Each xFacInfo In xColFacilities.Values
                If xFacInfo.OwnerID = OwnerID Then
                    dr = dtFacility.NewRow()
                    dr("OWNER_ID") = xFacInfo.OwnerID
                    dr("FACILITY_ID") = xFacInfo.ID
                    dr("CAP_PARTICIPANT") = xFacInfo.CapStatus
                    dr("FACILITYNAME") = xFacInfo.Name
                    xAddressInfo = pOwn2.Facilities.FacilityAddresses.Retrieve(xFacInfo.AddressID)
                    dr("ADDRESS") = xAddressInfo.AddressLine1
                    dr("CITY") = xAddressInfo.City
                    sList.Add(xFacInfo.ID, dr)
                End If
            Next

            For i = 0 To sList.Count - 1
                dtFacility.Rows.Add(CType(sList.GetByIndex(i), DataRow))
            Next

            Return dtFacility
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Friend Sub CurrentOwnerInformation(Optional ByVal data As ListView.ListViewItemCollection = Nothing)
        Dim dtDrOwn1Facility As DataRow
        'Dim rConsumer As New RegConsumer
        Dim dtOwn1Facilities As DataTable
        Dim lstViewItem As ListViewItem
        Dim nCompliantFacs As Integer = 0
        Dim nNonCompliantFacs As Integer = 0
        'Dim i As Integer = 0

        Try
            'dtOwn1Facilities = PopulateFacilities(pOwn.ID)
            'dtOwn1Facilities = pOwn.Facilities.EntityTablewithAddressDetails(pOwn.ID).Tables(0)
            If data Is Nothing Then
                dtOwn1Facilities = PopulateFacilities(nOwnerID)
            Else
                dtOwn1Facilities = New DataTable

                With dtOwn1Facilities.Columns
                    .Add("FAC", ("").GetType)

                    .Add("FACILITY_ID", (0).GetType)
                    .Add("FACILITYNAME", ("").GetType)
                    .Add("ADDRESS", ("").GetType)
                    .Add("CITY", ("").GetType)
                    .Add("CAP_PARTICIPANT", (0).GetType)
                End With


                For Each thisrow As ListViewItem In data
                    Dim nr As DataRow = dtOwn1Facilities.NewRow()
                    nr("FAC") = thisrow.SubItems.Item(0).Text
                    nr("FACILITY_ID") = Convert.ToInt32(thisrow.SubItems.Item(1).Text)
                    nr("FACILITYNAME") = thisrow.SubItems.Item(2).Text
                    nr("ADDRESS") = thisrow.SubItems.Item(3).Text
                    nr("CITY") = thisrow.SubItems.Item(4).Text
                    nr("CAP_PARTICIPANT") = IIf(thisrow.SubItems.Item(5).Text.ToUpper = "0", False, True)

                    dtOwn1Facilities.Rows.Add(nr)
                Next

            End If


            dtGlobal = dtOwn1Facilities
            lstViewFacLeft.Items.Clear()


            For Each dtDrOwn1Facility In dtOwn1Facilities.Select(String.Empty, sortIndex)
                Me.lstViewFacLeft.Items.Add(New ListViewItem(New String() {"", dtDrOwn1Facility("FACILITY_ID"), dtDrOwn1Facility("FACILITYNAME"), dtDrOwn1Facility("ADDRESS"), dtDrOwn1Facility("CITY"), dtDrOwn1Facility("CAP_PARTICIPANT")}))
                'dtGlobal.ImportRow(dtDrOwn1Facility)
            Next
            Me.lblNoOfFacilitiesOwnerValue.Text = lstViewFacLeft.Items.Count
            For Each lstViewItem In lstViewFacLeft.Items
                If lstViewItem.SubItems(5).Text = 1 Then
                    nCompliantFacs = nCompliantFacs + 1
                ElseIf lstViewItem.SubItems(5).Text = 0 Then
                    nNonCompliantFacs = nNonCompliantFacs + 1
                End If
            Next
            'MR - Starts
            'If i = lstViewFacLeft.Items.Count And i <> 0 Then
            '    Me.lblCAPParticipantValueOwner1.Text = " Yes"
            'ElseIf i < lstViewFacLeft.Items.Count And i <> 0 Then
            '    Me.lblCAPParticipantValueOwner1.Text = " Partial"
            'Else
            ' Me.lblCAPParticipantValueOwner1.Text = "None"
            'End If
            'MR - Ends
            If nCompliantFacs > 0 And nNonCompliantFacs = 0 Then
                Me.lblCAPParticipantValueOwner1.Text = "Compliant"
            ElseIf nCompliantFacs > 0 And nNonCompliantFacs > 0 Then
                Me.lblCAPParticipantValueOwner1.Text = "Partial Compliant"
            ElseIf nCompliantFacs = 0 And nNonCompliantFacs >= 0 Then
                Me.lblCAPParticipantValueOwner1.Text = "Non Compliant"
            End If

            dtOwn1Facilities.Dispose()
            data = Nothing

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub cmbOwnerRight_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOwnerRight.SelectedIndexChanged
        Dim dtOwn2Facilities As DataTable
        Dim dtDrOwn2Facility As DataRow
        'Dim rConsumer As New RegConsumer
        Dim lstViewItem As ListViewItem
        Dim lstViewFacilityDetailItem As ListViewItem
        Dim drFac1 As DataRow
        'Dim i As Integer = 0
        Dim nCompliantFacs As Integer = 0
        Dim nNonCompliantFacs As Integer = 0
        Dim strOwnerId As String
        Try
            If bolLoading Then Exit Sub

            Me.lstViewFacRight.Items.Clear()



            'If lstViewFacRight.Items.Count > 0 Then
            '    For Each lstViewItem In lstViewFacRight.Items
            '        If Not lstViewItem.Checked Then
            '            lstViewItem.Remove()
            '        Else
            '            i = lstViewItem.Index
            '            For Each lstViewFacilityDetailItem In Me.lstViewFacilityDetail.Items
            '                If i = lstViewFacilityDetailItem.SubItems(5).Text Then
            '                    i = lstViewFacilityDetailItem.Index
            '                    lstViewFacilityDetail.Items(i).SubItems(0).Text = cmbOwnerRight.Text
            '                End If
            '            Next
            '        End If
            '    Next
            'End If


            If cmbOwnerRight.SelectedIndex <> -1 Then
                '                strOwnerId = cmbOwnerRight.SelectedValue.ToString
                '               If strOwnerId <> "InfoRepository.LookupProperty" Then

                'dtOwn2Facilities = rConsumer.getFacilitiesOwnershipTransfer(cmbOwnerRight.SelectedValue)
                ' To Reload the Selected Owner's Facility List
                If lstViewFacLeft.Items.Count <> dtGlobal.Rows.Count Then
                    Me.lstViewFacLeft.Items.Clear()
                    For Each drFac1 In dtGlobal.Rows
                        Me.lstViewFacLeft.Items.Add(New ListViewItem(New String() {"", drFac1("FACILITY_ID"), drFac1("FACILITYNAME"), drFac1("ADDRESS"), drFac1("CITY"), drFac1("CAP_PARTICIPANT")}))
                    Next
                End If


                dtOwn2Facilities = PopulateFacilities(CLng(cmbOwnerRight.SelectedValue))
                '          End If
            End If
            If Not IsNothing(dtOwn2Facilities) Then
                For Each dtDrOwn2Facility In dtOwn2Facilities.Rows

                    Me.lstViewFacRight.Items.Add(New ListViewItem(New String() {"", dtDrOwn2Facility("FACILITY_ID"), dtDrOwn2Facility("FACILITYNAME"), dtDrOwn2Facility("ADDRESS"), dtDrOwn2Facility("CITY"), dtDrOwn2Facility("CAP_PARTICIPANT")}))
                Next
            End If

            Me.lblNoOfFacilitiesOwnerPotentialValue.Text = lstViewFacRight.Items.Count

            'i = 0
            If cmbOwnerRight.SelectedIndex <> -1 Then
                For Each lstViewItem In lstViewFacRight.Items
                    If lstViewItem.SubItems(5).Text = 1 Then
                        'i = i + 1
                        nCompliantFacs = nCompliantFacs + 1
                    ElseIf lstViewItem.SubItems(5).Text = 0 Then
                        nNonCompliantFacs = nNonCompliantFacs + 1
                    End If
                Next
                'MR - Starts
                'If i = lstViewFacRight.Items.Count And i <> 0 Then
                '    Me.lblCAPParticipantValueOwner2.Text = " Yes"
                'ElseIf i < lstViewFacRight.Items.Count And i <> 0 Then
                '    Me.lblCAPParticipantValueOwner2.Text = " Partial"
                'Else
                'Me.lblCAPParticipantValueOwner2.Text = "None"
                'End If
                'MR - Ends
                If nCompliantFacs > 0 And nNonCompliantFacs = 0 Then
                    Me.lblCAPParticipantValueOwner2.Text = "Compliant"
                ElseIf nCompliantFacs > 0 And nNonCompliantFacs > 0 Then
                    Me.lblCAPParticipantValueOwner2.Text = "Partial Compliant"
                ElseIf nCompliantFacs = 0 And nNonCompliantFacs >= 0 Then
                    Me.lblCAPParticipantValueOwner2.Text = "Non Compliant"

                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot load Facilities: " + ex.Message, ex))
            MyErr.ShowDialog()
            'MsgBox("Cannot load Facilities: " + ex.Message)
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim msgResult As MsgBoxResult
        Try
            msgResult = MsgBox("Do you want to close Transfer Ownership ? ", MsgBoxStyle.YesNo, "Transfer Ownership")
            If msgResult = MsgBoxResult.Yes Then
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lstViewFacLeft_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lstViewFacLeft.ItemCheck
        Dim lstViewItem As ListViewItem
        Try

            If e.CurrentValue = CheckState.Unchecked Then
                lstViewItem = CType((lstViewFacLeft.Items.Item(e.Index)), ListViewItem)
                Me.lstViewFacilityDetail.Items.Add(New ListViewItem(New String() {Me.txtBoxOwnerLeft.Text, lstViewItem.SubItems(1).Text, lstViewItem.SubItems(2).Text, "", "", e.Index}))
                'lstViewFacLeft.Items.Item(e.Index).Checked = False

            ElseIf e.CurrentValue = CheckState.Checked Then
                For Each lstViewItem In lstViewFacilityDetail.Items
                    If e.Index = lstViewItem.SubItems(5).Text Then
                        Dim i As Integer = lstViewItem.Index
                        Me.lstViewFacilityDetail.Items.RemoveAt(i)
                    End If
                Next

            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Display Facility Details: " + ex.Message, ex))
            MyErr.ShowDialog()
            'MsgBox("Cannot Display Facility Details: " + ex.Message)
        End Try
    End Sub
    Private Sub lstViewFacRight_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lstViewFacRight.ItemCheck
        Dim lstViewItem As ListViewItem
        Dim lstViewFacilityDetailItem As ListViewItem
        Dim i As Integer
        'lstViewFacilityDetail.Items.Clear()
        Try
            lstViewItem = CType((lstViewFacRight.Items.Item(e.Index)), ListViewItem)
            For Each lstViewFacilityDetailItem In lstViewFacilityDetail.Items
                If lstViewItem.SubItems(1).Text = lstViewFacilityDetailItem.SubItems(1).Text Then
                    i = lstViewFacilityDetailItem.Index
                    lstViewFacilityDetail.Items.RemoveAt(i)
                End If
            Next

            If e.CurrentValue = CheckState.Unchecked Then


                Me.lstViewFacilityDetail.Items.Add(New ListViewItem(New String() {Me.cmbOwnerRight.Text, lstViewItem.SubItems(1).Text, lstViewItem.SubItems(2).Text, "", "", e.Index}))
                'lstViewFacLeft.Items.Item(e.Index).Checked = False


            ElseIf e.CurrentValue = CheckState.Checked Then
                For Each lstViewItem In lstViewFacilityDetail.Items
                    If e.Index = lstViewItem.SubItems(5).Text Then
                        i = lstViewItem.Index
                        Me.lstViewFacilityDetail.Items.RemoveAt(i)
                    End If
                Next



            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Display Facility Details: " + ex.Message, ex))
            MyErr.ShowDialog()
            'MsgBox("Cannot Display Facility Details: " + ex.Message)
        End Try
    End Sub

    Private Sub TransferOwnership_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim arrList As ArrayList
        'Dim rConsumer As New RegConsumer
        Dim dtOwnerName As DataTable

        Try
            'arrList = rConsumer.getOwnerNames(CLng(pOwn.ID))
            dtOwnerName = pOwn.PopulateOwnerNameAndOwnerID(CLng(pOwn.ID))

            Dim dr As DataRow
            dr = dtOwnerName.NewRow
            dr("o_id") = 0
            dr("o_name") = " - Please choose an owner"
            dtOwnerName.Rows.InsertAt(dr, 0)

            'Me.cmbOwnerRight.DataSource = arrList
            bolLoading = True
            Me.cmbOwnerRight.DataSource = dtOwnerName
            Me.cmbOwnerRight.DisplayMember = "o_name"
            Me.cmbOwnerRight.ValueMember = "o_id"

            CurrentOwnerInformation()

            cmbOwnerRight.SelectedIndex = 0
            'cmbOwnerRight.SelectedIndex = -1
            'If cmbOwnerRight.SelectedIndex <> -1 Then
            '    cmbOwnerRight.SelectedIndex = -1
            'End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Owners: " + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub

    Private Sub lstViewFacLeft_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lstViewFacLeft.ColumnClick

        Me.sortIndex = Me.lstViewFacLeft.Columns(e.Column).Text.ToUpper.Replace(" ID", "_ID").Replace(" ", String.Empty)

        Me.CurrentOwnerInformation(Me.lstViewFacLeft.Items)

    End Sub
End Class
