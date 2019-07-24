Public Class RemediationSystem
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.RemediationSystem.vb
    '   Provides the window for the Remediation System Add/Update.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        AB      05/05/05    Original class definition.
    '  
    '
    '  Mode Logic:
    '   If nActivityID > 0 and nSystemID > 0 then
    '       = Update Existing System on an Activity
    '
    '   If nActivityID > 0 and nSystemID = 0 then
    '       = Add New System to an Activity
    '
    '   If nActivityID = 0 and nSystemID > 0 then
    '       = Update Existing System with no association with an Activity
    '
    '   If nActivityID = 0 and nSystemID = 0 then
    '       = Add New System with no association with an Activity
    '
    '-------------------------------------------------------------------------------
    '
    Inherits System.Windows.Forms.Form
#Region " Local Variables "
    Friend Mode As Int16
    Friend nSystemID As Int64
    Friend nActivityID As Int64
    Friend CallingForm As Form

    Private bolLoading As Boolean

    Private WithEvents oLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
    Private oLustRemediation As New MUSTER.BusinessLogic.pLustRemediation
    Private returnVal As String = String.Empty

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
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents lblSerialNumber As System.Windows.Forms.Label
    Friend WithEvents lblManufacturerName As System.Windows.Forms.Label
    Friend WithEvents lblModelNumber As System.Windows.Forms.Label
    Friend WithEvents lblRemedySystemLocation As System.Windows.Forms.Label
    Friend WithEvents txtStripperSize As System.Windows.Forms.TextBox
    Friend WithEvents txtMotorSize As System.Windows.Forms.TextBox
    Friend WithEvents txtOWSSize As System.Windows.Forms.TextBox
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents lblUsedNew As System.Windows.Forms.Label
    Friend WithEvents lblAgeofComponents As System.Windows.Forms.Label
    Friend WithEvents dtPickPurchaseDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbOptionalEquipment1 As System.Windows.Forms.ComboBox
    Friend WithEvents lblOptionalEquipment As System.Windows.Forms.Label
    Friend WithEvents cmbOptionalEquipment2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbOptionalEquipment3 As System.Windows.Forms.ComboBox
    Friend WithEvents lblOther As System.Windows.Forms.Label
    Friend WithEvents txtOther As System.Windows.Forms.TextBox
    Friend WithEvents lblRefurbishedDate As System.Windows.Forms.Label
    Friend WithEvents lblMount As System.Windows.Forms.Label
    Friend WithEvents lblBuildingSize As System.Windows.Forms.Label
    Friend WithEvents txtBuildingSize As System.Windows.Forms.TextBox
    Friend WithEvents lblPurchaseDate As System.Windows.Forms.Label
    Friend WithEvents lblOwner As System.Windows.Forms.Label
    Friend WithEvents txtOwner As System.Windows.Forms.TextBox
    Friend WithEvents cmbOwnedLeased As System.Windows.Forms.ComboBox
    Friend WithEvents lblOwnedLeased As System.Windows.Forms.Label
    Friend WithEvents cmbMount As System.Windows.Forms.ComboBox
    Friend WithEvents dtPickRefurbishedDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtOWSModelNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtOWSSerialNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtOWSManufacturerName As System.Windows.Forms.TextBox
    Friend WithEvents lblStripper As System.Windows.Forms.Label
    Friend WithEvents lblMotor As System.Windows.Forms.Label
    Friend WithEvents lblOWS As System.Windows.Forms.Label
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents cmbOWSUsedNew As System.Windows.Forms.ComboBox
    Friend WithEvents txtOWSAgeofComponents As System.Windows.Forms.TextBox
    Friend WithEvents lblSize As System.Windows.Forms.Label
    Friend WithEvents txtMotorAgeofComponents As System.Windows.Forms.TextBox
    Friend WithEvents txtMotorModelNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtMotorSerialNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtMotorManufacturerName As System.Windows.Forms.TextBox
    Friend WithEvents cmbMotorUsedNew As System.Windows.Forms.ComboBox
    Friend WithEvents txtStripperManufacturerName As System.Windows.Forms.TextBox
    Friend WithEvents txtStripperSerialNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtStripperModelNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtStripperAgeofComponents As System.Windows.Forms.TextBox
    Friend WithEvents cmbStripperUsedNew As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbVacPump1Seal As System.Windows.Forms.ComboBox
    Friend WithEvents txtVacPump2ManufacturerName As System.Windows.Forms.TextBox
    Friend WithEvents txtVacPump2SerialNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtVacPump2ModelNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtVacPump2AgeofComponents As System.Windows.Forms.TextBox
    Friend WithEvents cmbVacPump2UsedNew As System.Windows.Forms.ComboBox
    Friend WithEvents txtVacPump1ManufacturerName As System.Windows.Forms.TextBox
    Friend WithEvents txtVacPump1SerialNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtVacPump1ModelNumber As System.Windows.Forms.TextBox
    Friend WithEvents txtVacPump1AgeofComponents As System.Windows.Forms.TextBox
    Friend WithEvents txtVacPump2Size As System.Windows.Forms.TextBox
    Friend WithEvents txtVacPump1Size As System.Windows.Forms.TextBox
    Friend WithEvents cmbVacPump1UsedNew As System.Windows.Forms.ComboBox
    Friend WithEvents cmbVacPump2Seal As System.Windows.Forms.ComboBox
    Friend WithEvents lblVacPump2 As System.Windows.Forms.Label
    Friend WithEvents lblVacpump1 As System.Windows.Forms.Label
    Friend WithEvents lblSequentialNumber As System.Windows.Forms.Label
    Friend WithEvents lbPreviousLocations As System.Windows.Forms.ListBox
    Friend WithEvents txtManufacturer As System.Windows.Forms.TextBox
    Friend WithEvents Descriptionasd As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblModelNumber = New System.Windows.Forms.Label
        Me.txtOWSModelNumber = New System.Windows.Forms.TextBox
        Me.lblSerialNumber = New System.Windows.Forms.Label
        Me.txtOWSSerialNumber = New System.Windows.Forms.TextBox
        Me.lblManufacturerName = New System.Windows.Forms.Label
        Me.txtOWSManufacturerName = New System.Windows.Forms.TextBox
        Me.lblStripper = New System.Windows.Forms.Label
        Me.txtStripperSize = New System.Windows.Forms.TextBox
        Me.lblMotor = New System.Windows.Forms.Label
        Me.txtMotorSize = New System.Windows.Forms.TextBox
        Me.lblOWS = New System.Windows.Forms.Label
        Me.txtOWSSize = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.lblOther = New System.Windows.Forms.Label
        Me.txtOther = New System.Windows.Forms.TextBox
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.lblType = New System.Windows.Forms.Label
        Me.lblStartDate = New System.Windows.Forms.Label
        Me.lblRemedySystemLocation = New System.Windows.Forms.Label
        Me.lblDescription = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.cmbOWSUsedNew = New System.Windows.Forms.ComboBox
        Me.lblUsedNew = New System.Windows.Forms.Label
        Me.lblAgeofComponents = New System.Windows.Forms.Label
        Me.txtOWSAgeofComponents = New System.Windows.Forms.TextBox
        Me.cmbOptionalEquipment1 = New System.Windows.Forms.ComboBox
        Me.lblOptionalEquipment = New System.Windows.Forms.Label
        Me.cmbOptionalEquipment2 = New System.Windows.Forms.ComboBox
        Me.cmbOptionalEquipment3 = New System.Windows.Forms.ComboBox
        Me.lblRefurbishedDate = New System.Windows.Forms.Label
        Me.lblMount = New System.Windows.Forms.Label
        Me.lblBuildingSize = New System.Windows.Forms.Label
        Me.txtBuildingSize = New System.Windows.Forms.TextBox
        Me.lblPurchaseDate = New System.Windows.Forms.Label
        Me.lblOwner = New System.Windows.Forms.Label
        Me.txtOwner = New System.Windows.Forms.TextBox
        Me.cmbOwnedLeased = New System.Windows.Forms.ComboBox
        Me.lblOwnedLeased = New System.Windows.Forms.Label
        Me.cmbMount = New System.Windows.Forms.ComboBox
        Me.dtPickPurchaseDate = New System.Windows.Forms.DateTimePicker
        Me.dtPickRefurbishedDate = New System.Windows.Forms.DateTimePicker
        Me.dtPickStartDate = New System.Windows.Forms.DateTimePicker
        Me.lblSize = New System.Windows.Forms.Label
        Me.txtMotorAgeofComponents = New System.Windows.Forms.TextBox
        Me.txtMotorModelNumber = New System.Windows.Forms.TextBox
        Me.txtMotorSerialNumber = New System.Windows.Forms.TextBox
        Me.txtMotorManufacturerName = New System.Windows.Forms.TextBox
        Me.cmbMotorUsedNew = New System.Windows.Forms.ComboBox
        Me.txtStripperManufacturerName = New System.Windows.Forms.TextBox
        Me.txtStripperSerialNumber = New System.Windows.Forms.TextBox
        Me.txtStripperModelNumber = New System.Windows.Forms.TextBox
        Me.txtStripperAgeofComponents = New System.Windows.Forms.TextBox
        Me.cmbStripperUsedNew = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbVacPump1Seal = New System.Windows.Forms.ComboBox
        Me.txtVacPump2ManufacturerName = New System.Windows.Forms.TextBox
        Me.txtVacPump2SerialNumber = New System.Windows.Forms.TextBox
        Me.txtVacPump2ModelNumber = New System.Windows.Forms.TextBox
        Me.txtVacPump2AgeofComponents = New System.Windows.Forms.TextBox
        Me.cmbVacPump2UsedNew = New System.Windows.Forms.ComboBox
        Me.txtVacPump1ManufacturerName = New System.Windows.Forms.TextBox
        Me.txtVacPump1SerialNumber = New System.Windows.Forms.TextBox
        Me.txtVacPump1ModelNumber = New System.Windows.Forms.TextBox
        Me.txtVacPump1AgeofComponents = New System.Windows.Forms.TextBox
        Me.txtVacPump2Size = New System.Windows.Forms.TextBox
        Me.txtVacPump1Size = New System.Windows.Forms.TextBox
        Me.cmbVacPump1UsedNew = New System.Windows.Forms.ComboBox
        Me.cmbVacPump2Seal = New System.Windows.Forms.ComboBox
        Me.lblVacPump2 = New System.Windows.Forms.Label
        Me.lblVacpump1 = New System.Windows.Forms.Label
        Me.lblSequentialNumber = New System.Windows.Forms.Label
        Me.lbPreviousLocations = New System.Windows.Forms.ListBox
        Me.txtManufacturer = New System.Windows.Forms.TextBox
        Me.Descriptionasd = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblModelNumber
        '
        Me.lblModelNumber.Location = New System.Drawing.Point(464, 224)
        Me.lblModelNumber.Name = "lblModelNumber"
        Me.lblModelNumber.Size = New System.Drawing.Size(96, 16)
        Me.lblModelNumber.TabIndex = 263
        Me.lblModelNumber.Text = "Model Number:"
        Me.lblModelNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOWSModelNumber
        '
        Me.txtOWSModelNumber.Location = New System.Drawing.Point(480, 248)
        Me.txtOWSModelNumber.Name = "txtOWSModelNumber"
        Me.txtOWSModelNumber.Size = New System.Drawing.Size(96, 20)
        Me.txtOWSModelNumber.TabIndex = 13
        Me.txtOWSModelNumber.Text = ""
        '
        'lblSerialNumber
        '
        Me.lblSerialNumber.Location = New System.Drawing.Point(336, 224)
        Me.lblSerialNumber.Name = "lblSerialNumber"
        Me.lblSerialNumber.Size = New System.Drawing.Size(88, 16)
        Me.lblSerialNumber.TabIndex = 261
        Me.lblSerialNumber.Text = "Serial Number:"
        Me.lblSerialNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOWSSerialNumber
        '
        Me.txtOWSSerialNumber.Location = New System.Drawing.Point(344, 248)
        Me.txtOWSSerialNumber.Name = "txtOWSSerialNumber"
        Me.txtOWSSerialNumber.Size = New System.Drawing.Size(128, 20)
        Me.txtOWSSerialNumber.TabIndex = 12
        Me.txtOWSSerialNumber.Text = ""
        '
        'lblManufacturerName
        '
        Me.lblManufacturerName.Location = New System.Drawing.Point(200, 224)
        Me.lblManufacturerName.Name = "lblManufacturerName"
        Me.lblManufacturerName.Size = New System.Drawing.Size(112, 16)
        Me.lblManufacturerName.TabIndex = 259
        Me.lblManufacturerName.Text = "Manufacturer Name:"
        Me.lblManufacturerName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOWSManufacturerName
        '
        Me.txtOWSManufacturerName.Location = New System.Drawing.Point(208, 248)
        Me.txtOWSManufacturerName.Name = "txtOWSManufacturerName"
        Me.txtOWSManufacturerName.Size = New System.Drawing.Size(128, 20)
        Me.txtOWSManufacturerName.TabIndex = 11
        Me.txtOWSManufacturerName.Text = ""
        '
        'lblStripper
        '
        Me.lblStripper.Location = New System.Drawing.Point(40, 312)
        Me.lblStripper.Name = "lblStripper"
        Me.lblStripper.Size = New System.Drawing.Size(72, 16)
        Me.lblStripper.TabIndex = 257
        Me.lblStripper.Text = "Stripper:"
        Me.lblStripper.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtStripperSize
        '
        Me.txtStripperSize.Location = New System.Drawing.Point(120, 312)
        Me.txtStripperSize.Name = "txtStripperSize"
        Me.txtStripperSize.Size = New System.Drawing.Size(80, 20)
        Me.txtStripperSize.TabIndex = 22
        Me.txtStripperSize.Text = ""
        '
        'lblMotor
        '
        Me.lblMotor.Location = New System.Drawing.Point(40, 280)
        Me.lblMotor.Name = "lblMotor"
        Me.lblMotor.Size = New System.Drawing.Size(72, 16)
        Me.lblMotor.TabIndex = 255
        Me.lblMotor.Text = "Motor:"
        Me.lblMotor.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMotorSize
        '
        Me.txtMotorSize.Location = New System.Drawing.Point(120, 280)
        Me.txtMotorSize.Name = "txtMotorSize"
        Me.txtMotorSize.Size = New System.Drawing.Size(80, 20)
        Me.txtMotorSize.TabIndex = 16
        Me.txtMotorSize.Text = ""
        '
        'lblOWS
        '
        Me.lblOWS.Location = New System.Drawing.Point(48, 248)
        Me.lblOWS.Name = "lblOWS"
        Me.lblOWS.Size = New System.Drawing.Size(64, 16)
        Me.lblOWS.TabIndex = 253
        Me.lblOWS.Text = "OWS:"
        Me.lblOWS.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOWSSize
        '
        Me.txtOWSSize.Location = New System.Drawing.Point(120, 248)
        Me.txtOWSSize.Name = "txtOWSSize"
        Me.txtOWSSize.Size = New System.Drawing.Size(80, 20)
        Me.txtOWSSize.TabIndex = 10
        Me.txtOWSSize.Text = ""
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(456, 512)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 23)
        Me.btnCancel.TabIndex = 47
        Me.btnCancel.Text = "&Cancel"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(344, 512)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(104, 23)
        Me.btnOK.TabIndex = 46
        Me.btnOK.Text = "&Save"
        '
        'lblOther
        '
        Me.lblOther.Location = New System.Drawing.Point(64, 456)
        Me.lblOther.Name = "lblOther"
        Me.lblOther.Size = New System.Drawing.Size(48, 16)
        Me.lblOther.TabIndex = 249
        Me.lblOther.Text = "Other:"
        Me.lblOther.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOther
        '
        Me.txtOther.Location = New System.Drawing.Point(120, 456)
        Me.txtOther.Multiline = True
        Me.txtOther.Name = "txtOther"
        Me.txtOther.Size = New System.Drawing.Size(768, 40)
        Me.txtOther.TabIndex = 45
        Me.txtOther.Text = ""
        '
        'cmbType
        '
        Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbType.Location = New System.Drawing.Point(120, 128)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(144, 21)
        Me.cmbType.TabIndex = 3
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(48, 128)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(64, 16)
        Me.lblType.TabIndex = 246
        Me.lblType.Text = "Type:"
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblStartDate
        '
        Me.lblStartDate.Location = New System.Drawing.Point(8, 40)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(104, 16)
        Me.lblStartDate.TabIndex = 265
        Me.lblStartDate.Text = "System Start Date:"
        Me.lblStartDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRemedySystemLocation
        '
        Me.lblRemedySystemLocation.Location = New System.Drawing.Point(368, 8)
        Me.lblRemedySystemLocation.Name = "lblRemedySystemLocation"
        Me.lblRemedySystemLocation.Size = New System.Drawing.Size(152, 21)
        Me.lblRemedySystemLocation.TabIndex = 266
        Me.lblRemedySystemLocation.Text = "System Previous Location(s)"
        Me.lblRemedySystemLocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDescription
        '
        Me.lblDescription.Location = New System.Drawing.Point(0, 72)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(120, 16)
        Me.lblDescription.TabIndex = 270
        Me.lblDescription.Text = "Manufacturer:"
        Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(120, 99)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(296, 20)
        Me.txtDescription.TabIndex = 2
        Me.txtDescription.Text = ""
        '
        'cmbOWSUsedNew
        '
        Me.cmbOWSUsedNew.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOWSUsedNew.Location = New System.Drawing.Point(584, 248)
        Me.cmbOWSUsedNew.Name = "cmbOWSUsedNew"
        Me.cmbOWSUsedNew.Size = New System.Drawing.Size(88, 21)
        Me.cmbOWSUsedNew.TabIndex = 14
        '
        'lblUsedNew
        '
        Me.lblUsedNew.Location = New System.Drawing.Point(576, 224)
        Me.lblUsedNew.Name = "lblUsedNew"
        Me.lblUsedNew.Size = New System.Drawing.Size(72, 16)
        Me.lblUsedNew.TabIndex = 271
        Me.lblUsedNew.Text = "Used / New:"
        Me.lblUsedNew.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblAgeofComponents
        '
        Me.lblAgeofComponents.Location = New System.Drawing.Point(672, 224)
        Me.lblAgeofComponents.Name = "lblAgeofComponents"
        Me.lblAgeofComponents.Size = New System.Drawing.Size(112, 16)
        Me.lblAgeofComponents.TabIndex = 274
        Me.lblAgeofComponents.Text = "Age of Components:"
        Me.lblAgeofComponents.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOWSAgeofComponents
        '
        Me.txtOWSAgeofComponents.Location = New System.Drawing.Point(680, 248)
        Me.txtOWSAgeofComponents.Name = "txtOWSAgeofComponents"
        Me.txtOWSAgeofComponents.Size = New System.Drawing.Size(128, 20)
        Me.txtOWSAgeofComponents.TabIndex = 15
        Me.txtOWSAgeofComponents.Text = ""
        '
        'cmbOptionalEquipment1
        '
        Me.cmbOptionalEquipment1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOptionalEquipment1.Location = New System.Drawing.Point(120, 424)
        Me.cmbOptionalEquipment1.Name = "cmbOptionalEquipment1"
        Me.cmbOptionalEquipment1.Size = New System.Drawing.Size(224, 21)
        Me.cmbOptionalEquipment1.TabIndex = 42
        '
        'lblOptionalEquipment
        '
        Me.lblOptionalEquipment.Location = New System.Drawing.Point(0, 424)
        Me.lblOptionalEquipment.Name = "lblOptionalEquipment"
        Me.lblOptionalEquipment.Size = New System.Drawing.Size(112, 16)
        Me.lblOptionalEquipment.TabIndex = 275
        Me.lblOptionalEquipment.Text = "Optional Equipment:"
        Me.lblOptionalEquipment.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'cmbOptionalEquipment2
        '
        Me.cmbOptionalEquipment2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOptionalEquipment2.Location = New System.Drawing.Point(352, 424)
        Me.cmbOptionalEquipment2.Name = "cmbOptionalEquipment2"
        Me.cmbOptionalEquipment2.Size = New System.Drawing.Size(224, 21)
        Me.cmbOptionalEquipment2.TabIndex = 43
        '
        'cmbOptionalEquipment3
        '
        Me.cmbOptionalEquipment3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOptionalEquipment3.Location = New System.Drawing.Point(584, 424)
        Me.cmbOptionalEquipment3.Name = "cmbOptionalEquipment3"
        Me.cmbOptionalEquipment3.Size = New System.Drawing.Size(224, 21)
        Me.cmbOptionalEquipment3.TabIndex = 44
        '
        'lblRefurbishedDate
        '
        Me.lblRefurbishedDate.Location = New System.Drawing.Point(400, 160)
        Me.lblRefurbishedDate.Name = "lblRefurbishedDate"
        Me.lblRefurbishedDate.Size = New System.Drawing.Size(120, 16)
        Me.lblRefurbishedDate.TabIndex = 292
        Me.lblRefurbishedDate.Text = "Refurbished Date:"
        Me.lblRefurbishedDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblMount
        '
        Me.lblMount.Location = New System.Drawing.Point(456, 128)
        Me.lblMount.Name = "lblMount"
        Me.lblMount.Size = New System.Drawing.Size(64, 16)
        Me.lblMount.TabIndex = 290
        Me.lblMount.Text = "Mount:"
        Me.lblMount.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblBuildingSize
        '
        Me.lblBuildingSize.Location = New System.Drawing.Point(432, 96)
        Me.lblBuildingSize.Name = "lblBuildingSize"
        Me.lblBuildingSize.Size = New System.Drawing.Size(88, 16)
        Me.lblBuildingSize.TabIndex = 288
        Me.lblBuildingSize.Text = "Building Size:"
        Me.lblBuildingSize.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtBuildingSize
        '
        Me.txtBuildingSize.Location = New System.Drawing.Point(528, 99)
        Me.txtBuildingSize.Name = "txtBuildingSize"
        Me.txtBuildingSize.Size = New System.Drawing.Size(128, 20)
        Me.txtBuildingSize.TabIndex = 7
        Me.txtBuildingSize.Text = ""
        '
        'lblPurchaseDate
        '
        Me.lblPurchaseDate.Location = New System.Drawing.Point(408, 72)
        Me.lblPurchaseDate.Name = "lblPurchaseDate"
        Me.lblPurchaseDate.Size = New System.Drawing.Size(112, 16)
        Me.lblPurchaseDate.TabIndex = 286
        Me.lblPurchaseDate.Text = "Purchase Date:"
        Me.lblPurchaseDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOwner
        '
        Me.lblOwner.Location = New System.Drawing.Point(48, 192)
        Me.lblOwner.Name = "lblOwner"
        Me.lblOwner.Size = New System.Drawing.Size(64, 16)
        Me.lblOwner.TabIndex = 282
        Me.lblOwner.Text = "Owner:"
        Me.lblOwner.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtOwner
        '
        Me.txtOwner.Location = New System.Drawing.Point(120, 188)
        Me.txtOwner.Name = "txtOwner"
        Me.txtOwner.Size = New System.Drawing.Size(144, 20)
        Me.txtOwner.TabIndex = 5
        Me.txtOwner.Text = ""
        '
        'cmbOwnedLeased
        '
        Me.cmbOwnedLeased.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnedLeased.Location = New System.Drawing.Point(120, 158)
        Me.cmbOwnedLeased.Name = "cmbOwnedLeased"
        Me.cmbOwnedLeased.Size = New System.Drawing.Size(144, 21)
        Me.cmbOwnedLeased.TabIndex = 4
        '
        'lblOwnedLeased
        '
        Me.lblOwnedLeased.Location = New System.Drawing.Point(16, 160)
        Me.lblOwnedLeased.Name = "lblOwnedLeased"
        Me.lblOwnedLeased.Size = New System.Drawing.Size(96, 16)
        Me.lblOwnedLeased.TabIndex = 279
        Me.lblOwnedLeased.Text = "Owned / Leased:"
        Me.lblOwnedLeased.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'cmbMount
        '
        Me.cmbMount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMount.Location = New System.Drawing.Point(528, 128)
        Me.cmbMount.Name = "cmbMount"
        Me.cmbMount.Size = New System.Drawing.Size(128, 21)
        Me.cmbMount.TabIndex = 8
        '
        'dtPickPurchaseDate
        '
        Me.dtPickPurchaseDate.Checked = False
        Me.dtPickPurchaseDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPurchaseDate.Location = New System.Drawing.Point(528, 70)
        Me.dtPickPurchaseDate.Name = "dtPickPurchaseDate"
        Me.dtPickPurchaseDate.ShowCheckBox = True
        Me.dtPickPurchaseDate.Size = New System.Drawing.Size(104, 20)
        Me.dtPickPurchaseDate.TabIndex = 6
        '
        'dtPickRefurbishedDate
        '
        Me.dtPickRefurbishedDate.Checked = False
        Me.dtPickRefurbishedDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickRefurbishedDate.Location = New System.Drawing.Point(528, 159)
        Me.dtPickRefurbishedDate.Name = "dtPickRefurbishedDate"
        Me.dtPickRefurbishedDate.ShowCheckBox = True
        Me.dtPickRefurbishedDate.Size = New System.Drawing.Size(104, 20)
        Me.dtPickRefurbishedDate.TabIndex = 9
        '
        'dtPickStartDate
        '
        Me.dtPickStartDate.Checked = False
        Me.dtPickStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickStartDate.Location = New System.Drawing.Point(120, 41)
        Me.dtPickStartDate.Name = "dtPickStartDate"
        Me.dtPickStartDate.ShowCheckBox = True
        Me.dtPickStartDate.Size = New System.Drawing.Size(96, 20)
        Me.dtPickStartDate.TabIndex = 1
        '
        'lblSize
        '
        Me.lblSize.Location = New System.Drawing.Point(120, 224)
        Me.lblSize.Name = "lblSize"
        Me.lblSize.Size = New System.Drawing.Size(32, 16)
        Me.lblSize.TabIndex = 253
        Me.lblSize.Text = "Size:"
        Me.lblSize.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtMotorAgeofComponents
        '
        Me.txtMotorAgeofComponents.Location = New System.Drawing.Point(680, 280)
        Me.txtMotorAgeofComponents.Name = "txtMotorAgeofComponents"
        Me.txtMotorAgeofComponents.Size = New System.Drawing.Size(128, 20)
        Me.txtMotorAgeofComponents.TabIndex = 21
        Me.txtMotorAgeofComponents.Text = ""
        '
        'txtMotorModelNumber
        '
        Me.txtMotorModelNumber.Location = New System.Drawing.Point(480, 280)
        Me.txtMotorModelNumber.Name = "txtMotorModelNumber"
        Me.txtMotorModelNumber.Size = New System.Drawing.Size(96, 20)
        Me.txtMotorModelNumber.TabIndex = 19
        Me.txtMotorModelNumber.Text = ""
        '
        'txtMotorSerialNumber
        '
        Me.txtMotorSerialNumber.Location = New System.Drawing.Point(344, 280)
        Me.txtMotorSerialNumber.Name = "txtMotorSerialNumber"
        Me.txtMotorSerialNumber.Size = New System.Drawing.Size(128, 20)
        Me.txtMotorSerialNumber.TabIndex = 18
        Me.txtMotorSerialNumber.Text = ""
        '
        'txtMotorManufacturerName
        '
        Me.txtMotorManufacturerName.Location = New System.Drawing.Point(208, 280)
        Me.txtMotorManufacturerName.Name = "txtMotorManufacturerName"
        Me.txtMotorManufacturerName.Size = New System.Drawing.Size(128, 20)
        Me.txtMotorManufacturerName.TabIndex = 17
        Me.txtMotorManufacturerName.Text = ""
        '
        'cmbMotorUsedNew
        '
        Me.cmbMotorUsedNew.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbMotorUsedNew.Location = New System.Drawing.Point(584, 280)
        Me.cmbMotorUsedNew.Name = "cmbMotorUsedNew"
        Me.cmbMotorUsedNew.Size = New System.Drawing.Size(88, 21)
        Me.cmbMotorUsedNew.TabIndex = 20
        '
        'txtStripperManufacturerName
        '
        Me.txtStripperManufacturerName.Location = New System.Drawing.Point(208, 312)
        Me.txtStripperManufacturerName.Name = "txtStripperManufacturerName"
        Me.txtStripperManufacturerName.Size = New System.Drawing.Size(128, 20)
        Me.txtStripperManufacturerName.TabIndex = 23
        Me.txtStripperManufacturerName.Text = ""
        '
        'txtStripperSerialNumber
        '
        Me.txtStripperSerialNumber.Location = New System.Drawing.Point(344, 312)
        Me.txtStripperSerialNumber.Name = "txtStripperSerialNumber"
        Me.txtStripperSerialNumber.Size = New System.Drawing.Size(128, 20)
        Me.txtStripperSerialNumber.TabIndex = 24
        Me.txtStripperSerialNumber.Text = ""
        '
        'txtStripperModelNumber
        '
        Me.txtStripperModelNumber.Location = New System.Drawing.Point(480, 312)
        Me.txtStripperModelNumber.Name = "txtStripperModelNumber"
        Me.txtStripperModelNumber.Size = New System.Drawing.Size(96, 20)
        Me.txtStripperModelNumber.TabIndex = 25
        Me.txtStripperModelNumber.Text = ""
        '
        'txtStripperAgeofComponents
        '
        Me.txtStripperAgeofComponents.Location = New System.Drawing.Point(680, 312)
        Me.txtStripperAgeofComponents.Name = "txtStripperAgeofComponents"
        Me.txtStripperAgeofComponents.Size = New System.Drawing.Size(128, 20)
        Me.txtStripperAgeofComponents.TabIndex = 27
        Me.txtStripperAgeofComponents.Text = ""
        '
        'cmbStripperUsedNew
        '
        Me.cmbStripperUsedNew.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbStripperUsedNew.Location = New System.Drawing.Point(584, 312)
        Me.cmbStripperUsedNew.Name = "cmbStripperUsedNew"
        Me.cmbStripperUsedNew.Size = New System.Drawing.Size(88, 21)
        Me.cmbStripperUsedNew.TabIndex = 26
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(816, 224)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(32, 16)
        Me.Label1.TabIndex = 274
        Me.Label1.Text = "Seal"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbVacPump1Seal
        '
        Me.cmbVacPump1Seal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbVacPump1Seal.Location = New System.Drawing.Point(816, 344)
        Me.cmbVacPump1Seal.Name = "cmbVacPump1Seal"
        Me.cmbVacPump1Seal.Size = New System.Drawing.Size(72, 21)
        Me.cmbVacPump1Seal.TabIndex = 34
        '
        'txtVacPump2ManufacturerName
        '
        Me.txtVacPump2ManufacturerName.Location = New System.Drawing.Point(208, 376)
        Me.txtVacPump2ManufacturerName.Name = "txtVacPump2ManufacturerName"
        Me.txtVacPump2ManufacturerName.Size = New System.Drawing.Size(128, 20)
        Me.txtVacPump2ManufacturerName.TabIndex = 36
        Me.txtVacPump2ManufacturerName.Text = ""
        '
        'txtVacPump2SerialNumber
        '
        Me.txtVacPump2SerialNumber.Location = New System.Drawing.Point(344, 376)
        Me.txtVacPump2SerialNumber.Name = "txtVacPump2SerialNumber"
        Me.txtVacPump2SerialNumber.Size = New System.Drawing.Size(128, 20)
        Me.txtVacPump2SerialNumber.TabIndex = 37
        Me.txtVacPump2SerialNumber.Text = ""
        '
        'txtVacPump2ModelNumber
        '
        Me.txtVacPump2ModelNumber.Location = New System.Drawing.Point(480, 376)
        Me.txtVacPump2ModelNumber.Name = "txtVacPump2ModelNumber"
        Me.txtVacPump2ModelNumber.Size = New System.Drawing.Size(96, 20)
        Me.txtVacPump2ModelNumber.TabIndex = 38
        Me.txtVacPump2ModelNumber.Text = ""
        '
        'txtVacPump2AgeofComponents
        '
        Me.txtVacPump2AgeofComponents.Location = New System.Drawing.Point(680, 376)
        Me.txtVacPump2AgeofComponents.Name = "txtVacPump2AgeofComponents"
        Me.txtVacPump2AgeofComponents.Size = New System.Drawing.Size(128, 20)
        Me.txtVacPump2AgeofComponents.TabIndex = 40
        Me.txtVacPump2AgeofComponents.Text = ""
        '
        'cmbVacPump2UsedNew
        '
        Me.cmbVacPump2UsedNew.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbVacPump2UsedNew.Location = New System.Drawing.Point(584, 376)
        Me.cmbVacPump2UsedNew.Name = "cmbVacPump2UsedNew"
        Me.cmbVacPump2UsedNew.Size = New System.Drawing.Size(88, 21)
        Me.cmbVacPump2UsedNew.TabIndex = 39
        '
        'txtVacPump1ManufacturerName
        '
        Me.txtVacPump1ManufacturerName.Location = New System.Drawing.Point(208, 344)
        Me.txtVacPump1ManufacturerName.Name = "txtVacPump1ManufacturerName"
        Me.txtVacPump1ManufacturerName.Size = New System.Drawing.Size(128, 20)
        Me.txtVacPump1ManufacturerName.TabIndex = 29
        Me.txtVacPump1ManufacturerName.Text = ""
        '
        'txtVacPump1SerialNumber
        '
        Me.txtVacPump1SerialNumber.Location = New System.Drawing.Point(344, 344)
        Me.txtVacPump1SerialNumber.Name = "txtVacPump1SerialNumber"
        Me.txtVacPump1SerialNumber.Size = New System.Drawing.Size(128, 20)
        Me.txtVacPump1SerialNumber.TabIndex = 30
        Me.txtVacPump1SerialNumber.Text = ""
        '
        'txtVacPump1ModelNumber
        '
        Me.txtVacPump1ModelNumber.Location = New System.Drawing.Point(480, 344)
        Me.txtVacPump1ModelNumber.Name = "txtVacPump1ModelNumber"
        Me.txtVacPump1ModelNumber.Size = New System.Drawing.Size(96, 20)
        Me.txtVacPump1ModelNumber.TabIndex = 31
        Me.txtVacPump1ModelNumber.Text = ""
        '
        'txtVacPump1AgeofComponents
        '
        Me.txtVacPump1AgeofComponents.Location = New System.Drawing.Point(680, 344)
        Me.txtVacPump1AgeofComponents.Name = "txtVacPump1AgeofComponents"
        Me.txtVacPump1AgeofComponents.Size = New System.Drawing.Size(128, 20)
        Me.txtVacPump1AgeofComponents.TabIndex = 33
        Me.txtVacPump1AgeofComponents.Text = ""
        '
        'txtVacPump2Size
        '
        Me.txtVacPump2Size.Location = New System.Drawing.Point(120, 376)
        Me.txtVacPump2Size.Name = "txtVacPump2Size"
        Me.txtVacPump2Size.Size = New System.Drawing.Size(80, 20)
        Me.txtVacPump2Size.TabIndex = 35
        Me.txtVacPump2Size.Text = ""
        '
        'txtVacPump1Size
        '
        Me.txtVacPump1Size.Location = New System.Drawing.Point(120, 344)
        Me.txtVacPump1Size.Name = "txtVacPump1Size"
        Me.txtVacPump1Size.Size = New System.Drawing.Size(80, 20)
        Me.txtVacPump1Size.TabIndex = 28
        Me.txtVacPump1Size.Text = ""
        '
        'cmbVacPump1UsedNew
        '
        Me.cmbVacPump1UsedNew.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbVacPump1UsedNew.Location = New System.Drawing.Point(584, 344)
        Me.cmbVacPump1UsedNew.Name = "cmbVacPump1UsedNew"
        Me.cmbVacPump1UsedNew.Size = New System.Drawing.Size(88, 21)
        Me.cmbVacPump1UsedNew.TabIndex = 32
        '
        'cmbVacPump2Seal
        '
        Me.cmbVacPump2Seal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbVacPump2Seal.Location = New System.Drawing.Point(816, 376)
        Me.cmbVacPump2Seal.Name = "cmbVacPump2Seal"
        Me.cmbVacPump2Seal.Size = New System.Drawing.Size(72, 21)
        Me.cmbVacPump2Seal.TabIndex = 41
        '
        'lblVacPump2
        '
        Me.lblVacPump2.Location = New System.Drawing.Point(40, 376)
        Me.lblVacPump2.Name = "lblVacPump2"
        Me.lblVacPump2.Size = New System.Drawing.Size(72, 16)
        Me.lblVacPump2.TabIndex = 257
        Me.lblVacPump2.Text = "Vac Pump 2:"
        Me.lblVacPump2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVacpump1
        '
        Me.lblVacpump1.Location = New System.Drawing.Point(40, 344)
        Me.lblVacpump1.Name = "lblVacpump1"
        Me.lblVacpump1.Size = New System.Drawing.Size(72, 16)
        Me.lblVacpump1.TabIndex = 255
        Me.lblVacpump1.Text = "Vac Pump 1:"
        Me.lblVacpump1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSequentialNumber
        '
        Me.lblSequentialNumber.Location = New System.Drawing.Point(24, 8)
        Me.lblSequentialNumber.Name = "lblSequentialNumber"
        Me.lblSequentialNumber.Size = New System.Drawing.Size(232, 24)
        Me.lblSequentialNumber.TabIndex = 293
        Me.lblSequentialNumber.Text = "Sequence number"
        '
        'lbPreviousLocations
        '
        Me.lbPreviousLocations.Location = New System.Drawing.Point(528, 8)
        Me.lbPreviousLocations.Name = "lbPreviousLocations"
        Me.lbPreviousLocations.Size = New System.Drawing.Size(360, 56)
        Me.lbPreviousLocations.TabIndex = 295
        '
        'txtManufacturer
        '
        Me.txtManufacturer.Location = New System.Drawing.Point(120, 70)
        Me.txtManufacturer.Name = "txtManufacturer"
        Me.txtManufacturer.Size = New System.Drawing.Size(296, 20)
        Me.txtManufacturer.TabIndex = 296
        Me.txtManufacturer.Text = ""
        '
        'Descriptionasd
        '
        Me.Descriptionasd.Location = New System.Drawing.Point(0, 96)
        Me.Descriptionasd.Name = "Descriptionasd"
        Me.Descriptionasd.Size = New System.Drawing.Size(120, 16)
        Me.Descriptionasd.TabIndex = 297
        Me.Descriptionasd.Text = "Description:"
        Me.Descriptionasd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'RemediationSystem
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(904, 550)
        Me.Controls.Add(Me.txtManufacturer)
        Me.Controls.Add(Me.Descriptionasd)
        Me.Controls.Add(Me.lbPreviousLocations)
        Me.Controls.Add(Me.lblSequentialNumber)
        Me.Controls.Add(Me.dtPickStartDate)
        Me.Controls.Add(Me.dtPickRefurbishedDate)
        Me.Controls.Add(Me.dtPickPurchaseDate)
        Me.Controls.Add(Me.txtBuildingSize)
        Me.Controls.Add(Me.txtOwner)
        Me.Controls.Add(Me.txtOWSAgeofComponents)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.txtOWSModelNumber)
        Me.Controls.Add(Me.txtOWSSerialNumber)
        Me.Controls.Add(Me.txtOWSManufacturerName)
        Me.Controls.Add(Me.txtStripperSize)
        Me.Controls.Add(Me.txtMotorSize)
        Me.Controls.Add(Me.txtOWSSize)
        Me.Controls.Add(Me.txtOther)
        Me.Controls.Add(Me.txtMotorAgeofComponents)
        Me.Controls.Add(Me.txtMotorModelNumber)
        Me.Controls.Add(Me.txtMotorSerialNumber)
        Me.Controls.Add(Me.txtMotorManufacturerName)
        Me.Controls.Add(Me.txtStripperManufacturerName)
        Me.Controls.Add(Me.txtStripperSerialNumber)
        Me.Controls.Add(Me.txtStripperModelNumber)
        Me.Controls.Add(Me.txtStripperAgeofComponents)
        Me.Controls.Add(Me.txtVacPump2ManufacturerName)
        Me.Controls.Add(Me.txtVacPump2SerialNumber)
        Me.Controls.Add(Me.txtVacPump2ModelNumber)
        Me.Controls.Add(Me.txtVacPump2AgeofComponents)
        Me.Controls.Add(Me.txtVacPump1ManufacturerName)
        Me.Controls.Add(Me.txtVacPump1SerialNumber)
        Me.Controls.Add(Me.txtVacPump1ModelNumber)
        Me.Controls.Add(Me.txtVacPump1AgeofComponents)
        Me.Controls.Add(Me.txtVacPump2Size)
        Me.Controls.Add(Me.txtVacPump1Size)
        Me.Controls.Add(Me.cmbMount)
        Me.Controls.Add(Me.lblRefurbishedDate)
        Me.Controls.Add(Me.lblMount)
        Me.Controls.Add(Me.lblBuildingSize)
        Me.Controls.Add(Me.lblOwner)
        Me.Controls.Add(Me.cmbOwnedLeased)
        Me.Controls.Add(Me.lblOwnedLeased)
        Me.Controls.Add(Me.cmbOptionalEquipment3)
        Me.Controls.Add(Me.cmbOptionalEquipment2)
        Me.Controls.Add(Me.cmbOptionalEquipment1)
        Me.Controls.Add(Me.lblOptionalEquipment)
        Me.Controls.Add(Me.lblAgeofComponents)
        Me.Controls.Add(Me.cmbOWSUsedNew)
        Me.Controls.Add(Me.lblUsedNew)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.lblRemedySystemLocation)
        Me.Controls.Add(Me.lblStartDate)
        Me.Controls.Add(Me.lblModelNumber)
        Me.Controls.Add(Me.lblSerialNumber)
        Me.Controls.Add(Me.lblManufacturerName)
        Me.Controls.Add(Me.lblStripper)
        Me.Controls.Add(Me.lblMotor)
        Me.Controls.Add(Me.lblOWS)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.lblOther)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.lblPurchaseDate)
        Me.Controls.Add(Me.lblSize)
        Me.Controls.Add(Me.cmbMotorUsedNew)
        Me.Controls.Add(Me.cmbStripperUsedNew)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbVacPump1Seal)
        Me.Controls.Add(Me.cmbVacPump2UsedNew)
        Me.Controls.Add(Me.cmbVacPump1UsedNew)
        Me.Controls.Add(Me.cmbVacPump2Seal)
        Me.Controls.Add(Me.lblVacPump2)
        Me.Controls.Add(Me.lblVacpump1)
        Me.Name = "RemediationSystem"
        Me.Text = "Add / Select Remediation System"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Populate Routines "
    Private Sub LoadDatatables()
        Dim rData As DataRow

        Try
            bolLoading = True
            Dim dtPrevLocations As DataTable = oLustRemediation.SystemPreviousLocations(oLustRemediation.SystemSequence)

            Dim dtOptEquip1 As DataTable = oLustRemediation.PopulateRemediationOptionalEquipment

            Dim dtOptEquip2 As DataTable = oLustRemediation.PopulateRemediationOptionalEquipment

            Dim dtOptEquip3 As DataTable = oLustRemediation.PopulateRemediationOptionalEquipment

            Dim dtNewUsed1 As DataTable = oLustRemediation.PopulateRemediationNewUsed
            rData = dtNewUsed1.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtNewUsed1.Rows.InsertAt(rData, 0)

            Dim dtNewUsed2 As DataTable = oLustRemediation.PopulateRemediationNewUsed
            rData = dtNewUsed2.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtNewUsed2.Rows.InsertAt(rData, 0)

            Dim dtNewUsed3 As DataTable = oLustRemediation.PopulateRemediationNewUsed
            rData = dtNewUsed3.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtNewUsed3.Rows.InsertAt(rData, 0)

            Dim dtNewUsed4 As DataTable = oLustRemediation.PopulateRemediationNewUsed
            rData = dtNewUsed4.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtNewUsed4.Rows.InsertAt(rData, 0)

            Dim dtNewUsed5 As DataTable = oLustRemediation.PopulateRemediationNewUsed
            rData = dtNewUsed5.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtNewUsed5.Rows.InsertAt(rData, 0)

            Dim dtPumpSeal1 As DataTable = oLustRemediation.PopulateRemediationPumpSeal
            rData = dtPumpSeal1.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtPumpSeal1.Rows.InsertAt(rData, 0)

            Dim dtPumpSeal2 As DataTable = oLustRemediation.PopulateRemediationPumpSeal
            rData = dtPumpSeal2.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtPumpSeal2.Rows.InsertAt(rData, 0)

            Dim dtMountType As DataTable = oLustRemediation.PopulateRemediationMountType
            rData = dtMountType.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtMountType.Rows.InsertAt(rData, 0)

            Dim dtType As DataTable = oLustRemediation.PopulateRemediationType
            rData = dtType.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtType.Rows.InsertAt(rData, 0)

            Dim dtOwnLease As DataTable = oLustRemediation.PopulateRemediationOwnedLeased
            rData = dtOwnLease.NewRow
            rData.Item("PROPERTY_NAME") = " "
            rData.Item("PROPERTY_ID") = 0
            dtOwnLease.Rows.InsertAt(rData, 0)

            lbPreviousLocations.DataSource = dtPrevLocations
            If lbPreviousLocations.Items.Count > 0 Then
                lbPreviousLocations.DisplayMember = "Full_Name"
                lbPreviousLocations.ValueMember = "Facility_ID"
            End If

            cmbOptionalEquipment1.DataSource = dtOptEquip1
            cmbOptionalEquipment1.DisplayMember = "PROPERTY_NAME"
            cmbOptionalEquipment1.ValueMember = "PROPERTY_ID"

            cmbOptionalEquipment2.DataSource = dtOptEquip2
            cmbOptionalEquipment2.DisplayMember = "PROPERTY_NAME"
            cmbOptionalEquipment2.ValueMember = "PROPERTY_ID"

            cmbOptionalEquipment3.DataSource = dtOptEquip3
            cmbOptionalEquipment3.DisplayMember = "PROPERTY_NAME"
            cmbOptionalEquipment3.ValueMember = "PROPERTY_ID"

            cmbOWSUsedNew.Items.Clear()
            cmbOWSUsedNew.DataSource = dtNewUsed1
            cmbOWSUsedNew.DisplayMember = "PROPERTY_NAME"
            cmbOWSUsedNew.ValueMember = "PROPERTY_ID"

            cmbMotorUsedNew.DataSource = dtNewUsed2
            cmbMotorUsedNew.DisplayMember = "PROPERTY_NAME"
            cmbMotorUsedNew.ValueMember = "PROPERTY_ID"

            cmbStripperUsedNew.DataSource = dtNewUsed3
            cmbStripperUsedNew.DisplayMember = "PROPERTY_NAME"
            cmbStripperUsedNew.ValueMember = "PROPERTY_ID"

            cmbVacPump1UsedNew.DataSource = dtNewUsed4
            cmbVacPump1UsedNew.DisplayMember = "PROPERTY_NAME"
            cmbVacPump1UsedNew.ValueMember = "PROPERTY_ID"

            cmbVacPump2UsedNew.DataSource = dtNewUsed5
            cmbVacPump2UsedNew.DisplayMember = "PROPERTY_NAME"
            cmbVacPump2UsedNew.ValueMember = "PROPERTY_ID"

            cmbVacPump1Seal.DataSource = dtPumpSeal1
            cmbVacPump1Seal.DisplayMember = "PROPERTY_NAME"
            cmbVacPump1Seal.ValueMember = "PROPERTY_ID"

            cmbVacPump2Seal.DataSource = dtPumpSeal2
            cmbVacPump2Seal.DisplayMember = "PROPERTY_NAME"
            cmbVacPump2Seal.ValueMember = "PROPERTY_ID"

            cmbMount.DataSource = dtMountType
            cmbMount.DisplayMember = "PROPERTY_NAME"
            cmbMount.ValueMember = "PROPERTY_ID"

            cmbOwnedLeased.DataSource = dtOwnLease
            cmbOwnedLeased.DisplayMember = "PROPERTY_NAME"
            cmbOwnedLeased.ValueMember = "PROPERTY_ID"

            cmbType.DataSource = dtType
            cmbType.DisplayMember = "PROPERTY_NAME"
            cmbType.ValueMember = "PROPERTY_ID"
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Remediation System DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try

    End Sub


    Private Sub LoadRemediationSystem()
        Dim tmpDate As Date

        Try
            bolLoading = True

            If Mode = 0 And nActivityID = 0 And nSystemID = 0 Then
                UIUtilsGen.SetDatePickerValue(dtPickStartDate, Now.Date)
                UIUtilsGen.ToggleDateFormat(dtPickStartDate)
                FillDateobjectValues(oLustRemediation.DateInUse, dtPickStartDate.Text)
            ElseIf Mode = 0 And nActivityID > 0 Then
                UIUtilsGen.SetDatePickerValue(dtPickStartDate, oLustActivity.Started)
                oLustRemediation.DateInUse = oLustActivity.Started.Date
            Else
                UIUtilsGen.SetDatePickerValue(dtPickStartDate, oLustRemediation.DateInUse)
            End If

            txtDescription.Text = oLustRemediation.Description
            txtManufacturer.Text = oLustRemediation.Manufacturer
            UIUtilsGen.SetDatePickerValue(dtPickPurchaseDate, oLustRemediation.PurchaseDate)
            SetDropDown(cmbType, oLustRemediation.RemSysType)

            txtBuildingSize.Text = oLustRemediation.BuildingSize
            SetDropDown(cmbOwnedLeased, oLustRemediation.Owned)
            If nSystemID = 0 Then
                cmbMount.SelectedIndex = 1
            Else
                SetDropDown(cmbMount, oLustRemediation.MountType)
            End If


            txtOwner.Text = oLustRemediation.Owner
            UIUtilsGen.SetDatePickerValue(dtPickRefurbishedDate, oLustRemediation.RefurbDate)

            txtOWSSize.Text = oLustRemediation.OWSSize
            txtOWSSerialNumber.Text = oLustRemediation.OWSSerialNumber
            txtOWSModelNumber.Text = oLustRemediation.OWSModelNumber
            txtOWSManufacturerName.Text = oLustRemediation.OWSManName
            txtOWSAgeofComponents.Text = oLustRemediation.OWSAgeofComp
            SetDropDown(cmbOWSUsedNew, oLustRemediation.OWSNewUsed)

            txtMotorSize.Text = oLustRemediation.MotorSize
            txtMotorSerialNumber.Text = oLustRemediation.MotorSerialNumber
            txtMotorModelNumber.Text = oLustRemediation.MotorModelNumber
            txtMotorManufacturerName.Text = oLustRemediation.MotorManName
            txtMotorAgeofComponents.Text = oLustRemediation.MotorAgeofComp
            SetDropDown(cmbMotorUsedNew, oLustRemediation.MotorNewUsed)

            txtStripperSize.Text = oLustRemediation.StripperSize
            txtStripperSerialNumber.Text = oLustRemediation.StripperSerialNumber
            txtStripperModelNumber.Text = oLustRemediation.StripperModelNumber
            txtStripperManufacturerName.Text = oLustRemediation.StripperManName
            txtStripperAgeofComponents.Text = oLustRemediation.StripperAgeofComp
            SetDropDown(cmbStripperUsedNew, oLustRemediation.StripperNewUsed)

            txtVacPump1Size.Text = oLustRemediation.VacPump1Size
            txtVacPump1SerialNumber.Text = oLustRemediation.VacPump1SerialNumber
            txtVacPump1ModelNumber.Text = oLustRemediation.VacPump1ModelNumber
            txtVacPump1ManufacturerName.Text = oLustRemediation.VacPump1ManName
            txtVacPump1AgeofComponents.Text = oLustRemediation.VacPump1AgeofComp
            SetDropDown(cmbVacPump1UsedNew, oLustRemediation.VacPump1NewUsed)
            SetDropDown(cmbVacPump1Seal, oLustRemediation.VacPump1Seal)

            txtVacPump2Size.Text = oLustRemediation.VacPump2Size
            txtVacPump2SerialNumber.Text = oLustRemediation.VacPump2SerialNumber
            txtVacPump2ModelNumber.Text = oLustRemediation.VacPump2ModelNumber
            txtVacPump2ManufacturerName.Text = oLustRemediation.VacPump2ManName
            txtVacPump2AgeofComponents.Text = oLustRemediation.VacPump2AgeofComp
            SetDropDown(cmbVacPump2UsedNew, oLustRemediation.VacPump2NewUsed)
            SetDropDown(cmbVacPump2Seal, oLustRemediation.VacPump2Seal)

            SetDropDown(cmbOptionalEquipment1, oLustRemediation.Option1)
            SetDropDown(cmbOptionalEquipment2, oLustRemediation.Option2)
            SetDropDown(cmbOptionalEquipment3, oLustRemediation.Option3)

            txtOther.Text = oLustRemediation.Notes

            bolLoading = False

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Remediation System " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try


    End Sub

    Private Sub SetDropDown(ByRef ctlDropDown As ComboBox, ByVal Value As Int64)
        Try

            ctlDropDown.SelectedValue = CInt(Value)
        Catch ex As Exception
            ctlDropDown.SelectedIndex = 0
        End Try
    End Sub

    Private Sub FillDateobjectValues(ByRef currentObj As Object, ByVal value As String)

        If value.Length > 0 And value <> "__/__/____" Then
            currentObj = CType(value, Date)
        Else
            currentObj = "#12:00:00AM#"
        End If
    End Sub

#End Region

#Region " Page Event Routines "

    Private Sub RemediationSystem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If nActivityID > 0 Then
            oLustActivity.Retrieve(nActivityID)
        End If
        oLustRemediation.Retrieve(nSystemID)

        Select Case Mode
            Case 0 ' New
                'Clone the remediation system for use with this activity....
                oLustRemediation.SystemDeclaration = 0
                oLustRemediation.ID = 0
                lblSequentialNumber.Text = "Sequence:  New System"
            Case 1 ' Existing
                lblSequentialNumber.Text = "Sequence:  " & oLustRemediation.SystemSequence & " - " & oLustRemediation.SystemDeclaration
            Case 3 ' Read Only
                btnOK.Visible = False
                btnCancel.Text = "Exit"
                lblSequentialNumber.Text = "Sequence:  " & oLustRemediation.SystemSequence & " - " & oLustRemediation.SystemDeclaration
        End Select

        LoadDatatables()
        LoadRemediationSystem()

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        'Perform Add/Update Here
        Dim tmpDate As Date
        Try
            If Me.cmbType.SelectedValue = 0 Then
                MsgBox("System Type Required.")
                Exit Sub
            End If
            'If oLustRemediation.DateInUse = tmpDate Then
            '    MsgBox("Start Date Required.")
            '    Exit Sub
            'End If

            If oLustRemediation.ID <= 0 Then
                oLustRemediation.CreatedBy = MusterContainer.AppUser.ID
            Else
                oLustRemediation.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oLustRemediation.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If Mode = 0 And nActivityID > 0 Then
                oLustActivity.RemSystemID = oLustRemediation.ID
                oLustActivity.ModifiedBy = MusterContainer.AppUser.ID
                oLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            End If
            MsgBox("Remediation System Saved")
            If Not CallingForm Is Nothing Then
                CallingForm.Tag = "1"
            End If
            Me.Close()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Save Remediation System " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        Me.Close()

    End Sub

    Private Sub dtPickStartDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickStartDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtPickStartDate)
        FillDateobjectValues(oLustRemediation.DateInUse, dtPickStartDate.Text)

    End Sub

    Private Sub txtDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescription.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.Description = txtDescription.Text

    End Sub

    Private Sub dtPickPurchaseDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickPurchaseDate.ValueChanged
        If bolLoading Then Exit Sub

        UIUtilsGen.ToggleDateFormat(dtPickPurchaseDate)
        FillDateobjectValues(oLustRemediation.PurchaseDate, dtPickPurchaseDate.Text)
    End Sub

    Private Sub cmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.RemSysType = cmbType.SelectedValue
    End Sub

    Private Sub txtBuildingSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuildingSize.TextChanged
        If bolLoading Then Exit Sub
        'If IsNumeric(txtBuildingSize.Text) Then
        oLustRemediation.BuildingSize = txtBuildingSize.Text
        'ElseIf txtBuildingSize.Text = "" Then
        '    oLustRemediation.BuildingSize = 0
        'Else
        '    txtBuildingSize.Text = "0"
        '    oLustRemediation.BuildingSize = 0
        '    MsgBox("Only Numeric Values Are Valid")
        'End If

    End Sub

    Private Sub cmbOwnedLeased_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnedLeased.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustRemediation.Owned = cmbOwnedLeased.SelectedValue
    End Sub

    Private Sub cmbMount_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMount.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustRemediation.MountType = cmbMount.SelectedValue
    End Sub

    Private Sub txtOwner_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwner.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.Owner = txtOwner.Text
    End Sub

    Private Sub dtPickRefurbishedDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickRefurbishedDate.ValueChanged
        If bolLoading Then Exit Sub

        UIUtilsGen.ToggleDateFormat(dtPickRefurbishedDate)
        FillDateobjectValues(oLustRemediation.RefurbDate, dtPickRefurbishedDate.Text)

    End Sub

    Private Sub txtOWSSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOWSSize.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtOWSSize.Text) Then
        oLustRemediation.OWSSize = txtOWSSize.Text
        'ElseIf txtOWSSize.Text = "" Then
        '    oLustRemediation.OWSSize = 0
        'Else
        '    txtOWSSize.Text = "0"
        '    oLustRemediation.OWSSize = 0
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub txtOWSManufacturerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOWSManufacturerName.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.OWSManName = txtOWSManufacturerName.Text
    End Sub

    Private Sub txtOWSSerialNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOWSSerialNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.OWSSerialNumber = txtOWSSerialNumber.Text
    End Sub

    Private Sub txtOWSModelNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOWSModelNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.OWSModelNumber = txtOWSModelNumber.Text
    End Sub

    Private Sub cmbOWSUsedNew_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOWSUsedNew.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustRemediation.OWSNewUsed = cmbOWSUsedNew.SelectedValue
    End Sub

    Private Sub txtOWSAgeofComponents_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOWSAgeofComponents.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtOWSAgeofComponents.Text) Then
        oLustRemediation.OWSAgeofComp = txtOWSAgeofComponents.Text
        'ElseIf txtOWSAgeofComponents.Text = "" Then

        '    oLustRemediation.OWSAgeofComp = 0
        'Else
        '    oLustRemediation.OWSAgeofComp = 0
        '    txtOWSAgeofComponents.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub txtMotorSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMotorSize.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtMotorSize.Text) Then
        oLustRemediation.MotorSize = txtMotorSize.Text
        'ElseIf txtMotorSize.Text = "" Then
        '    oLustRemediation.MotorSize = 0
        'Else
        '    oLustRemediation.MotorSize = 0
        '    txtMotorSize.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub txtMotorManufacturerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMotorManufacturerName.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.MotorManName = txtMotorManufacturerName.Text
    End Sub

    Private Sub txtMotorSerialNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMotorSerialNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.MotorSerialNumber = txtMotorSerialNumber.Text
    End Sub

    Private Sub txtMotorModelNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMotorModelNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.MotorModelNumber = txtMotorModelNumber.Text
    End Sub

    Private Sub cmbMotorUsedNew_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMotorUsedNew.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oLustRemediation.MotorNewUsed = cmbMotorUsedNew.SelectedValue
    End Sub

    Private Sub txtMotorAgeofComponents_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMotorAgeofComponents.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtMotorAgeofComponents.Text) Then
        oLustRemediation.MotorAgeofComp = txtMotorAgeofComponents.Text
        'ElseIf txtMotorAgeofComponents.Text = "" Then

        '    oLustRemediation.MotorAgeofComp = 0
        'Else
        '    oLustRemediation.MotorAgeofComp = 0
        '    txtMotorAgeofComponents.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub txtStripperSize_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStripperSize.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtStripperSize.Text) Then
        oLustRemediation.StripperSize = txtStripperSize.Text
        'ElseIf txtStripperSize.Text = "" Then
        '    oLustRemediation.StripperSize = 0
        'Else
        '    oLustRemediation.StripperSize = 0
        '    txtStripperSize.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub txtStripperManufacturerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStripperManufacturerName.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.StripperManName = txtStripperManufacturerName.Text
    End Sub

    Private Sub txtStripperSerialNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStripperSerialNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.StripperSerialNumber = txtStripperSerialNumber.Text
    End Sub

    Private Sub txtStripperModelNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStripperModelNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.StripperModelNumber = txtStripperModelNumber.Text
    End Sub

    Private Sub cmbStripperUsedNew_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbStripperUsedNew.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.StripperNewUsed = cmbStripperUsedNew.SelectedValue
    End Sub

    Private Sub txtStripperAgeofComponents_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStripperAgeofComponents.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtStripperAgeofComponents.Text) Then
        oLustRemediation.StripperAgeofComp = txtStripperAgeofComponents.Text
        'ElseIf txtStripperAgeofComponents.Text = "" Then
        '    oLustRemediation.StripperAgeofComp = 0
        'Else
        '    oLustRemediation.StripperAgeofComp = 0
        '    txtStripperAgeofComponents.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub txtVacPump1Size_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump1Size.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtVacPump1Size.Text) Then
        oLustRemediation.VacPump1Size = txtVacPump1Size.Text
        'ElseIf txtVacPump1Size.Text = "" Then
        '    oLustRemediation.VacPump1Size = 0
        'Else
        '    oLustRemediation.VacPump1Size = 0
        '    txtVacPump1Size.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub txtVacPump1ManufacturerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump1ManufacturerName.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump1ManName = txtVacPump1ManufacturerName.Text
    End Sub

    Private Sub txtVacPump1SerialNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump1SerialNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump1SerialNumber = txtVacPump1SerialNumber.Text
    End Sub

    Private Sub txtVacPump1ModelNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump1ModelNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump1ModelNumber = txtVacPump1ModelNumber.Text
    End Sub

    Private Sub cmbVacPump1UsedNew_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbVacPump1UsedNew.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump1NewUsed = cmbVacPump1UsedNew.SelectedValue
    End Sub

    Private Sub txtVacPump1AgeofComponents_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump1AgeofComponents.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtVacPump1AgeofComponents.Text) Then
        oLustRemediation.VacPump1AgeofComp = txtVacPump1AgeofComponents.Text
        'ElseIf txtVacPump1AgeofComponents.Text = "" Then
        '    oLustRemediation.VacPump1AgeofComp = 0
        'Else
        '    oLustRemediation.VacPump1AgeofComp = 0
        '    txtVacPump1AgeofComponents.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub cmbVacPump1Seal_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbVacPump1Seal.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump1Seal = cmbVacPump1Seal.SelectedValue
    End Sub

    Private Sub txtVacPump2Size_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump2Size.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtVacPump2Size.Text) Then
        oLustRemediation.VacPump2Size = txtVacPump2Size.Text
        'ElseIf txtVacPump2Size.Text = "" Then
        '    oLustRemediation.VacPump1Size = 0
        'Else
        '    oLustRemediation.VacPump2Size = 0
        '    txtVacPump2Size.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub txtVacPump2ManufacturerName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump2ManufacturerName.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump2ManName = txtVacPump2ManufacturerName.Text
    End Sub

    Private Sub txtVacPump2SerialNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump2SerialNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump2SerialNumber = txtVacPump2SerialNumber.Text
    End Sub

    Private Sub txtVacPump2ModelNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump2ModelNumber.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump2ModelNumber = txtVacPump2ModelNumber.Text
    End Sub

    Private Sub cmbVacPump2UsedNew_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbVacPump2UsedNew.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump2NewUsed = cmbVacPump2UsedNew.SelectedValue
    End Sub

    Private Sub txtVacPump2AgeofComponents_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVacPump2AgeofComponents.TextChanged
        If bolLoading Then Exit Sub
        'P1 03/09/06 start
        'If IsNumeric(txtVacPump2AgeofComponents.Text) Then
        oLustRemediation.VacPump2AgeofComp = txtVacPump2AgeofComponents.Text
        'ElseIf txtVacPump2AgeofComponents.Text = "" Then
        '    oLustRemediation.VacPump2AgeofComp = 0
        'Else
        '    oLustRemediation.VacPump2AgeofComp = 0
        '    txtVacPump2AgeofComponents.Text = "0"
        '    MsgBox("Only Numeric Values Are Valid")
        'End If
        'P1 03/09/06 end
    End Sub

    Private Sub cmbVacPump2Seal_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbVacPump2Seal.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.VacPump2Seal = cmbVacPump2Seal.SelectedValue
    End Sub

    Private Sub cmbOptionalEquipment1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOptionalEquipment1.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.Option1 = cmbOptionalEquipment1.SelectedValue
    End Sub

    Private Sub cmbOptionalEquipment2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOptionalEquipment2.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.Option2 = cmbOptionalEquipment2.SelectedValue
    End Sub

    Private Sub cmbOptionalEquipment3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOptionalEquipment3.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustRemediation.Option3 = cmbOptionalEquipment3.SelectedValue
    End Sub

    Private Sub txtOther_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOther.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.Notes = txtOther.Text
    End Sub

    Private Sub txtManufacturer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtManufacturer.TextChanged
        If bolLoading Then Exit Sub

        oLustRemediation.Manufacturer = txtManufacturer.Text
    End Sub

#End Region

End Class
