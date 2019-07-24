Public Class Document
    Inherits System.Windows.Forms.Form


#Region " Local Variables "

    Friend CallingForm As Form
    Friend Mode As Int16
    Friend EventActivityID As Int64
    Friend EventOwnerID As Int64
    Friend EventDocumentID As Int64
    Friend TFStatus As Int16
    Friend StartDate As Date
    Friend ActivityName As String
    Friend IsTech As Boolean = True

    Private nCurrentDocType As Integer
    Private bolLoading As Boolean
    Private WithEvents oLustDocument As New MUSTER.BusinessLogic.pLustEventDocument
    Private WithEvents oComments As New MUSTER.BusinessLogic.pComments
    Private oLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
    Private oLustEvent As MUSTER.BusinessLogic.pLustEvent
    Private oOwner As New MUSTER.BusinessLogic.pOwner
    Private dtComments As DataTable
    Private returnVal As String = String.Empty

    Dim CantNFADocumentID As Int64
    Dim bolCantNFA As Boolean
    Dim contNFA As Boolean = False

#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(ByRef pLustEvent As MUSTER.BusinessLogic.pLustEvent)
        MyBase.New()

        oLustEvent = pLustEvent

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
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents cmbDocument As System.Windows.Forms.ComboBox
    Friend WithEvents lblDocument As System.Windows.Forms.Label
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents txtActivity As System.Windows.Forms.TextBox
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents txtStartDate As System.Windows.Forms.TextBox
    Friend WithEvents lblIssued As System.Windows.Forms.Label
    Friend WithEvents lblDue As System.Windows.Forms.Label
    Friend WithEvents lblReceived As System.Windows.Forms.Label
    Friend WithEvents lblExtension As System.Windows.Forms.Label
    Friend WithEvents lblClosed As System.Windows.Forms.Label
    Friend WithEvents lblToFinancial As System.Windows.Forms.Label
    Friend WithEvents dtIssued As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtDue As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtExtension As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtReceived As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtToFinancial As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtClosed As System.Windows.Forms.DateTimePicker
    Friend WithEvents gbRevision1 As System.Windows.Forms.GroupBox
    Friend WithEvents dtRevision1Due As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtRevision1Received As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblRevision1Received As System.Windows.Forms.Label
    Friend WithEvents lblRevision1Due As System.Windows.Forms.Label
    Friend WithEvents gbRevision2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtRevision2Received As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtRevision2Due As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.lblComments = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.cmbDocument = New System.Windows.Forms.ComboBox
        Me.lblDocument = New System.Windows.Forms.Label
        Me.lblActivity = New System.Windows.Forms.Label
        Me.txtActivity = New System.Windows.Forms.TextBox
        Me.lblStartDate = New System.Windows.Forms.Label
        Me.txtStartDate = New System.Windows.Forms.TextBox
        Me.lblIssued = New System.Windows.Forms.Label
        Me.lblDue = New System.Windows.Forms.Label
        Me.lblReceived = New System.Windows.Forms.Label
        Me.lblExtension = New System.Windows.Forms.Label
        Me.lblClosed = New System.Windows.Forms.Label
        Me.lblToFinancial = New System.Windows.Forms.Label
        Me.dtIssued = New System.Windows.Forms.DateTimePicker
        Me.dtDue = New System.Windows.Forms.DateTimePicker
        Me.dtExtension = New System.Windows.Forms.DateTimePicker
        Me.dtReceived = New System.Windows.Forms.DateTimePicker
        Me.dtToFinancial = New System.Windows.Forms.DateTimePicker
        Me.dtClosed = New System.Windows.Forms.DateTimePicker
        Me.gbRevision1 = New System.Windows.Forms.GroupBox
        Me.lblRevision1Received = New System.Windows.Forms.Label
        Me.lblRevision1Due = New System.Windows.Forms.Label
        Me.dtRevision1Received = New System.Windows.Forms.DateTimePicker
        Me.dtRevision1Due = New System.Windows.Forms.DateTimePicker
        Me.gbRevision2 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.dtRevision2Received = New System.Windows.Forms.DateTimePicker
        Me.dtRevision2Due = New System.Windows.Forms.DateTimePicker
        Me.gbRevision1.SuspendLayout()
        Me.gbRevision2.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(224, 304)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 23)
        Me.btnCancel.TabIndex = 14
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(112, 304)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(104, 23)
        Me.btnSave.TabIndex = 13
        Me.btnSave.Text = "Save"
        '
        'lblComments
        '
        Me.lblComments.Location = New System.Drawing.Point(24, 232)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(80, 16)
        Me.lblComments.TabIndex = 213
        Me.lblComments.Text = "Comments:"
        Me.lblComments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(112, 232)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(504, 64)
        Me.txtComments.TabIndex = 12
        Me.txtComments.Text = ""
        '
        'cmbDocument
        '
        Me.cmbDocument.Location = New System.Drawing.Point(112, 48)
        Me.cmbDocument.Name = "cmbDocument"
        Me.cmbDocument.Size = New System.Drawing.Size(312, 21)
        Me.cmbDocument.TabIndex = 1
        '
        'lblDocument
        '
        Me.lblDocument.Location = New System.Drawing.Point(32, 48)
        Me.lblDocument.Name = "lblDocument"
        Me.lblDocument.Size = New System.Drawing.Size(64, 16)
        Me.lblDocument.TabIndex = 208
        Me.lblDocument.Text = "Document:"
        Me.lblDocument.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblActivity
        '
        Me.lblActivity.Location = New System.Drawing.Point(40, 16)
        Me.lblActivity.Name = "lblActivity"
        Me.lblActivity.Size = New System.Drawing.Size(56, 16)
        Me.lblActivity.TabIndex = 220
        Me.lblActivity.Text = "Activity:"
        Me.lblActivity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtActivity
        '
        Me.txtActivity.Location = New System.Drawing.Point(112, 16)
        Me.txtActivity.Name = "txtActivity"
        Me.txtActivity.ReadOnly = True
        Me.txtActivity.Size = New System.Drawing.Size(312, 20)
        Me.txtActivity.TabIndex = 0
        Me.txtActivity.TabStop = False
        Me.txtActivity.Text = ""
        '
        'lblStartDate
        '
        Me.lblStartDate.Location = New System.Drawing.Point(432, 16)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(80, 16)
        Me.lblStartDate.TabIndex = 221
        Me.lblStartDate.Text = "Start Date: "
        Me.lblStartDate.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'txtStartDate
        '
        Me.txtStartDate.Location = New System.Drawing.Point(520, 16)
        Me.txtStartDate.Name = "txtStartDate"
        Me.txtStartDate.ReadOnly = True
        Me.txtStartDate.Size = New System.Drawing.Size(96, 20)
        Me.txtStartDate.TabIndex = 1
        Me.txtStartDate.TabStop = False
        Me.txtStartDate.Text = ""
        '
        'lblIssued
        '
        Me.lblIssued.Location = New System.Drawing.Point(40, 80)
        Me.lblIssued.Name = "lblIssued"
        Me.lblIssued.Size = New System.Drawing.Size(56, 16)
        Me.lblIssued.TabIndex = 224
        Me.lblIssued.Text = "Issued:"
        Me.lblIssued.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDue
        '
        Me.lblDue.Location = New System.Drawing.Point(40, 104)
        Me.lblDue.Name = "lblDue"
        Me.lblDue.Size = New System.Drawing.Size(56, 16)
        Me.lblDue.TabIndex = 226
        Me.lblDue.Text = "Due:"
        Me.lblDue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblReceived
        '
        Me.lblReceived.Location = New System.Drawing.Point(40, 152)
        Me.lblReceived.Name = "lblReceived"
        Me.lblReceived.Size = New System.Drawing.Size(56, 16)
        Me.lblReceived.TabIndex = 230
        Me.lblReceived.Text = "Received:"
        Me.lblReceived.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblExtension
        '
        Me.lblExtension.Location = New System.Drawing.Point(32, 128)
        Me.lblExtension.Name = "lblExtension"
        Me.lblExtension.Size = New System.Drawing.Size(64, 16)
        Me.lblExtension.TabIndex = 228
        Me.lblExtension.Text = "Extension:"
        Me.lblExtension.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblClosed
        '
        Me.lblClosed.Location = New System.Drawing.Point(40, 200)
        Me.lblClosed.Name = "lblClosed"
        Me.lblClosed.Size = New System.Drawing.Size(56, 16)
        Me.lblClosed.TabIndex = 234
        Me.lblClosed.Text = "Closed:"
        Me.lblClosed.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblToFinancial
        '
        Me.lblToFinancial.Location = New System.Drawing.Point(8, 176)
        Me.lblToFinancial.Name = "lblToFinancial"
        Me.lblToFinancial.Size = New System.Drawing.Size(88, 16)
        Me.lblToFinancial.TabIndex = 232
        Me.lblToFinancial.Text = "To Financial:"
        Me.lblToFinancial.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtIssued
        '
        Me.dtIssued.Checked = False
        Me.dtIssued.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtIssued.Location = New System.Drawing.Point(128, 80)
        Me.dtIssued.Name = "dtIssued"
        Me.dtIssued.Size = New System.Drawing.Size(88, 20)
        Me.dtIssued.TabIndex = 2
        '
        'dtDue
        '
        Me.dtDue.Checked = False
        Me.dtDue.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtDue.Location = New System.Drawing.Point(112, 104)
        Me.dtDue.Name = "dtDue"
        Me.dtDue.ShowCheckBox = True
        Me.dtDue.Size = New System.Drawing.Size(104, 20)
        Me.dtDue.TabIndex = 3
        '
        'dtExtension
        '
        Me.dtExtension.Checked = False
        Me.dtExtension.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtExtension.Location = New System.Drawing.Point(112, 128)
        Me.dtExtension.Name = "dtExtension"
        Me.dtExtension.ShowCheckBox = True
        Me.dtExtension.Size = New System.Drawing.Size(104, 20)
        Me.dtExtension.TabIndex = 4
        '
        'dtReceived
        '
        Me.dtReceived.Checked = False
        Me.dtReceived.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtReceived.Location = New System.Drawing.Point(112, 152)
        Me.dtReceived.Name = "dtReceived"
        Me.dtReceived.ShowCheckBox = True
        Me.dtReceived.Size = New System.Drawing.Size(104, 20)
        Me.dtReceived.TabIndex = 5
        '
        'dtToFinancial
        '
        Me.dtToFinancial.Checked = False
        Me.dtToFinancial.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtToFinancial.Location = New System.Drawing.Point(112, 176)
        Me.dtToFinancial.Name = "dtToFinancial"
        Me.dtToFinancial.ShowCheckBox = True
        Me.dtToFinancial.Size = New System.Drawing.Size(104, 20)
        Me.dtToFinancial.TabIndex = 6
        '
        'dtClosed
        '
        Me.dtClosed.Checked = False
        Me.dtClosed.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtClosed.Location = New System.Drawing.Point(112, 200)
        Me.dtClosed.Name = "dtClosed"
        Me.dtClosed.ShowCheckBox = True
        Me.dtClosed.Size = New System.Drawing.Size(104, 20)
        Me.dtClosed.TabIndex = 7
        '
        'gbRevision1
        '
        Me.gbRevision1.Controls.Add(Me.lblRevision1Received)
        Me.gbRevision1.Controls.Add(Me.lblRevision1Due)
        Me.gbRevision1.Controls.Add(Me.dtRevision1Received)
        Me.gbRevision1.Controls.Add(Me.dtRevision1Due)
        Me.gbRevision1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbRevision1.Location = New System.Drawing.Point(232, 80)
        Me.gbRevision1.Name = "gbRevision1"
        Me.gbRevision1.Size = New System.Drawing.Size(184, 88)
        Me.gbRevision1.TabIndex = 255
        Me.gbRevision1.TabStop = False
        Me.gbRevision1.Text = "Revision # 1"
        '
        'lblRevision1Received
        '
        Me.lblRevision1Received.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRevision1Received.Location = New System.Drawing.Point(8, 56)
        Me.lblRevision1Received.Name = "lblRevision1Received"
        Me.lblRevision1Received.Size = New System.Drawing.Size(56, 16)
        Me.lblRevision1Received.TabIndex = 253
        Me.lblRevision1Received.Text = "Received:"
        Me.lblRevision1Received.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblRevision1Due
        '
        Me.lblRevision1Due.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRevision1Due.Location = New System.Drawing.Point(8, 24)
        Me.lblRevision1Due.Name = "lblRevision1Due"
        Me.lblRevision1Due.Size = New System.Drawing.Size(32, 16)
        Me.lblRevision1Due.TabIndex = 252
        Me.lblRevision1Due.Text = "Due:"
        Me.lblRevision1Due.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dtRevision1Received
        '
        Me.dtRevision1Received.Checked = False
        Me.dtRevision1Received.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtRevision1Received.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRevision1Received.Location = New System.Drawing.Point(72, 56)
        Me.dtRevision1Received.Name = "dtRevision1Received"
        Me.dtRevision1Received.ShowCheckBox = True
        Me.dtRevision1Received.Size = New System.Drawing.Size(104, 20)
        Me.dtRevision1Received.TabIndex = 9
        '
        'dtRevision1Due
        '
        Me.dtRevision1Due.Checked = False
        Me.dtRevision1Due.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtRevision1Due.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRevision1Due.Location = New System.Drawing.Point(72, 24)
        Me.dtRevision1Due.Name = "dtRevision1Due"
        Me.dtRevision1Due.ShowCheckBox = True
        Me.dtRevision1Due.Size = New System.Drawing.Size(104, 20)
        Me.dtRevision1Due.TabIndex = 8
        '
        'gbRevision2
        '
        Me.gbRevision2.Controls.Add(Me.Label1)
        Me.gbRevision2.Controls.Add(Me.Label2)
        Me.gbRevision2.Controls.Add(Me.dtRevision2Received)
        Me.gbRevision2.Controls.Add(Me.dtRevision2Due)
        Me.gbRevision2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbRevision2.Location = New System.Drawing.Point(432, 80)
        Me.gbRevision2.Name = "gbRevision2"
        Me.gbRevision2.Size = New System.Drawing.Size(184, 88)
        Me.gbRevision2.TabIndex = 256
        Me.gbRevision2.TabStop = False
        Me.gbRevision2.Text = "Revision # 2"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 253
        Me.Label1.Text = "Received:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 16)
        Me.Label2.TabIndex = 252
        Me.Label2.Text = "Due:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dtRevision2Received
        '
        Me.dtRevision2Received.Checked = False
        Me.dtRevision2Received.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtRevision2Received.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRevision2Received.Location = New System.Drawing.Point(72, 56)
        Me.dtRevision2Received.Name = "dtRevision2Received"
        Me.dtRevision2Received.ShowCheckBox = True
        Me.dtRevision2Received.Size = New System.Drawing.Size(104, 20)
        Me.dtRevision2Received.TabIndex = 11
        '
        'dtRevision2Due
        '
        Me.dtRevision2Due.Checked = False
        Me.dtRevision2Due.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtRevision2Due.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRevision2Due.Location = New System.Drawing.Point(72, 24)
        Me.dtRevision2Due.Name = "dtRevision2Due"
        Me.dtRevision2Due.ShowCheckBox = True
        Me.dtRevision2Due.Size = New System.Drawing.Size(104, 20)
        Me.dtRevision2Due.TabIndex = 10
        '
        'Document
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(632, 350)
        Me.Controls.Add(Me.gbRevision2)
        Me.Controls.Add(Me.gbRevision1)
        Me.Controls.Add(Me.dtClosed)
        Me.Controls.Add(Me.dtToFinancial)
        Me.Controls.Add(Me.dtReceived)
        Me.Controls.Add(Me.dtExtension)
        Me.Controls.Add(Me.dtDue)
        Me.Controls.Add(Me.dtIssued)
        Me.Controls.Add(Me.lblClosed)
        Me.Controls.Add(Me.lblToFinancial)
        Me.Controls.Add(Me.lblReceived)
        Me.Controls.Add(Me.lblExtension)
        Me.Controls.Add(Me.lblDue)
        Me.Controls.Add(Me.lblIssued)
        Me.Controls.Add(Me.txtStartDate)
        Me.Controls.Add(Me.lblStartDate)
        Me.Controls.Add(Me.lblActivity)
        Me.Controls.Add(Me.txtActivity)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.lblComments)
        Me.Controls.Add(Me.txtComments)
        Me.Controls.Add(Me.cmbDocument)
        Me.Controls.Add(Me.lblDocument)
        Me.Name = "Document"
        Me.Text = "Document"
        Me.gbRevision1.ResumeLayout(False)
        Me.gbRevision2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Events "
    Private Sub Document_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim tmpDate As Date
        Dim oDocument As New MUSTER.BusinessLogic.pTecDoc

        oLustActivity.Retrieve(EventActivityID)
        oLustDocument.Retrieve(EventDocumentID)
        oOwner.Retrieve(EventOwnerID)

        'dtComments = oComments.GetByModule("Technical", 24, oLustDocument.ID)
        Dim lustDocID As Integer
        If oLustDocument.ID = 0 Then
            lustDocID = -1
        Else
            lustDocID = oLustDocument.ID
        End If
        dtComments = oComments.GetComments("Technical", 24, lustDocID).Tables(0)
        'If oLustDocument.ID = 0 Then
        '    dtComments.Rows.Clear()
        'End If

        If dtComments.Rows.Count > 0 Then
            oComments.Retrieve(dtComments.Rows(0)("COMMENT_ID"), dtComments.Rows(0)("USER ID"))
            If oComments.Deleted = False Then
                txtComments.Text = oComments.Comments
            Else
                txtComments.Text = String.Empty
            End If

        End If

        If TFStatus = 617 Or TFStatus = 620 Then
            dtToFinancial.Enabled = False
        End If
        txtActivity.Text = ActivityName
        txtStartDate.Text = IIf(StartDate = tmpDate, "", StartDate.Date)

        bolLoading = True

        PopulateLustDocuments()
        If Mode = 0 Then ' Add Mode
            oLustDocument.IssueDate = Now.Date
            oLustDocument.EventId = oLustActivity.EventID
            oLustDocument.AssocActivity = EventActivityID
            UIUtilsGen.SetDatePickerValue(dtClosed, tmpDate)
            UIUtilsGen.SetDatePickerValue(dtDue, tmpDate)
            UIUtilsGen.SetDatePickerValue(dtIssued, Now.Date)
            UIUtilsGen.SetDatePickerValue(dtReceived, tmpDate)
            UIUtilsGen.SetDatePickerValue(dtToFinancial, tmpDate)

            UIUtilsGen.SetDatePickerValue(dtExtension, tmpDate)
            UIUtilsGen.SetDatePickerValue(dtRevision1Due, tmpDate)
            UIUtilsGen.SetDatePickerValue(dtRevision1Received, tmpDate)
            UIUtilsGen.SetDatePickerValue(dtRevision2Due, tmpDate)
            UIUtilsGen.SetDatePickerValue(dtRevision2Received, tmpDate)
            dtExtension.Enabled = False
            gbRevision1.Enabled = False
            gbRevision2.Enabled = False
            cmbDocument.Enabled = True
            btnSave.Enabled = True
        Else 'Update Mode
            oDocument.Retrieve(oLustDocument.DocumentType)
            nCurrentDocType = oDocument.DocType
            SetDocumentType()
            UIUtilsGen.SetDatePickerValue(dtClosed, oLustDocument.DocClosedDate)
            UIUtilsGen.SetDatePickerValue(dtDue, oLustDocument.DueDate)
            UIUtilsGen.SetDatePickerValue(dtIssued, oLustDocument.IssueDate)
            UIUtilsGen.SetDatePickerValue(dtReceived, oLustDocument.DocRcvDate)
            UIUtilsGen.SetDatePickerValue(dtToFinancial, oLustDocument.DocFinancialDate)
            UIUtilsGen.SetDatePickerValue(dtExtension, oLustDocument.EXTENSIONDATE)
            UIUtilsGen.SetDatePickerValue(dtRevision1Due, oLustDocument.REV1EXTENSIONDATE)
            UIUtilsGen.SetDatePickerValue(dtRevision1Received, oLustDocument.REV1RECEIVEDDATE)
            UIUtilsGen.SetDatePickerValue(dtRevision2Due, oLustDocument.REV2EXTENSIONDATE)
            UIUtilsGen.SetDatePickerValue(dtRevision2Received, oLustDocument.REV2RECEIVEDDATE)
            cmbDocument.Enabled = False
            'dtDue.Enabled = False
            If Not MusterContainer.AppUser.HEAD_PM Then
                dtIssued.Enabled = False
            End If
            btnSave.Enabled = False
        End If

        If Not IsTech Then

            dtReceived.Enabled = False
            dtExtension.Enabled = False
            dtRevision1Due.Enabled = False
            dtRevision2Due.Enabled = False
            dtToFinancial.Enabled = True
            dtIssued.Enabled = False
            dtDue.Enabled = False
            dtClosed.Enabled = False

            dtRevision1Received.Enabled = False
            dtRevision2Received.Enabled = False
            txtComments.ReadOnly = True
            cmbDocument.Enabled = False


        End If

        bolLoading = False
    End Sub



    Private Sub dtToFinancial_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtToFinancial.ValueChanged
        If dtToFinancial.Checked Then
            dtClosed.Enabled = False
        End If

        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtToFinancial)
        FillDateobjectValues(oLustDocument.DocFinancialDate, dtToFinancial.Text)
        If oLustDocument.DocFinancialDate = "01/01/0001" Then
            dtClosed.Enabled = True
        End If

    End Sub

    Private Sub dtIssued_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtIssued.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtIssued)
        FillDateobjectValues(oLustDocument.IssueDate, dtIssued.Text)
    End Sub

    Private Sub dtDue_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtDue.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtDue)
        FillDateobjectValues(oLustDocument.DueDate, dtDue.Text)
    End Sub

    Private Sub dtExtension_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtExtension.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtExtension)
        FillDateobjectValues(oLustDocument.EXTENSIONDATE, dtExtension.Text)

    End Sub

    Private Sub dtReceived_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtReceived.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtReceived)
        FillDateobjectValues(oLustDocument.DocRcvDate, dtReceived.Text)

    End Sub

    Private Sub dtClosed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtClosed.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtClosed)
        FillDateobjectValues(oLustDocument.DocClosedDate, dtClosed.Text)

    End Sub

    Private Sub dtRevision1Due_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRevision1Due.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtRevision1Due)
        FillDateobjectValues(oLustDocument.REV1EXTENSIONDATE, dtRevision1Due.Text)
    End Sub

    Private Sub dtRevision2Due_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRevision2Due.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtRevision2Due)
        FillDateobjectValues(oLustDocument.REV2EXTENSIONDATE, dtRevision2Due.Text)

    End Sub

    Private Sub dtRevision1Received_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRevision1Received.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtRevision1Received)
        FillDateobjectValues(oLustDocument.REV1RECEIVEDDATE, dtRevision1Received.Text)
    End Sub

    Private Sub dtRevision2Received_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRevision2Received.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtRevision2Received)
        FillDateobjectValues(oLustDocument.REV2RECEIVEDDATE, dtRevision2Received.Text)
    End Sub

    Private Sub cmbDocument_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDocument.SelectedIndexChanged
        If bolLoading Then Exit Sub

        If cmbDocument.SelectedValue = 747 Or cmbDocument.SelectedValue = 784 Then
            ' Test for NFA Acceptance

        End If

        oLustDocument.DocumentType = cmbDocument.SelectedValue
        oLustDocument.DocClass = cmbDocument.SelectedValue
        SetDocumentType()

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub oLustDocument_LustEventChanged(ByVal bolValue As Boolean) Handles oLustDocument.LustEventChanged
        If bolValue = True Or txtComments.Text <> oComments.Comments Then
            btnSave.Enabled = True
        Else
            btnSave.Enabled = False
        End If
    End Sub



    Private Sub txtComments_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComments.LostFocus
        If Mode <> 0 And oComments.ID > 0 Then 'If Not Add Mode and a Comment already exists
            If txtComments.Text = String.Empty Then
                oComments.Deleted = True
                btnSave.Enabled = True
            ElseIf txtComments.Text <> oComments.Comments Then
                oComments.Deleted = False
                oComments.Comments = txtComments.Text
                btnSave.Enabled = True
            End If
        Else
            If txtComments.Text <> String.Empty Or oLustDocument.IsDirty Then
                btnSave.Enabled = True
            Else
                btnSave.Enabled = False
            End If
        End If
    End Sub

    Private Sub txtComments_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtComments.TextChanged
        btnSave.Enabled = True
    End Sub
#End Region

#Region " Populate Routines "
    Private Sub PopulateLustDocuments()
        'If bolLoading Then Exit Sub
        Try
            Dim dtLustActivityDocs As DataTable = oLustActivity.PopulateLustDocuments(TFStatus)
            If Not IsNothing(dtLustActivityDocs) Then
                cmbDocument.DataSource = dtLustActivityDocs
                cmbDocument.DisplayMember = "DocName"
                cmbDocument.ValueMember = "Document_ID"
            Else
                cmbDocument.DataSource = Nothing
            End If
            If Mode = 0 Then
                cmbDocument.SelectedIndex = -1
            Else
                cmbDocument.SelectedValue = oLustDocument.DocumentType
                If cmbDocument.SelectedValue <> oLustDocument.DocumentType Then
                    MsgBox("Document Invalid For This Activity Or Paytype")
                End If
            End If


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tec Activity Documents" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub FillDateobjectValues(ByRef currentObj As Object, ByVal value As String)

        If value.Length > 0 And value <> "__/__/____" Then
            currentObj = CType(value, Date).Date
        Else
            currentObj = "#12:00:00AM#"
        End If
    End Sub
#End Region


#Region " Comments "

    Private Sub ProcessComments()
        Try
            If Not IsTech Then
                txtComments.ReadOnly = False

                If txtComments.Text.Length > 1 Then
                    txtComments.Text = String.Format("{0}{1}", txtComments.Text.Replace(String.Format("{0} Sent To Financial by Financial Dept.", vbCrLf), String.Empty), String.Format("{0} Sent To Financial by Financial Dept.", vbCrLf))
                Else
                    txtComments.Text = String.Format("Sent To Financial by Financial Dept.")
                End If

                txtComments_LostFocus(Me, Nothing)

                txtComments.ReadOnly = True

            End If

            If Mode = 0 Then
                If txtComments.Text <> String.Empty Then
                    InsertComment()
                End If
            Else
                If oComments.ID > 0 Then
                    If oComments.IsDirty Then
                        oComments.ModifiedBy = MusterContainer.AppUser.ID
                        oComments.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    End If
                ElseIf txtComments.Text <> String.Empty Then
                    InsertComment()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub InsertComment()
        Dim oCommentInfo As New MUSTER.Info.CommentsInfo
        Try

            With oCommentInfo
                .CommentDate = Now.Date
                .Comments = txtComments.Text
                .CommentsScope = "External"
                .EntityID = oLustDocument.ID
                .EntityType = 24
                .ModuleName = "Technical"
                .UserID = MusterContainer.AppUser.ID

                If .ID <= 0 Then
                    .CreatedBy = MusterContainer.AppUser.ID
                Else
                    .ModifiedBy = MusterContainer.AppUser.ID
                End If
            End With

            oComments.Add(oCommentInfo)
            oComments.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
        Catch ex As Exception

        End Try

    End Sub

#End Region


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        ProcessSaveEvent()

    End Sub

    Private Sub ProcessSaveEvent()
        Dim bolNewDocument As Boolean = False
        Dim bolNewSentToFinancial As Boolean = False
        Dim bolNewClosedDate As Boolean = False
        Dim bolNewRecievedDate As Boolean = False
        Dim bolNewExtensionDate As Boolean = False
        Dim bolNewDueDate As Boolean = False
        Dim bolNewRevisionDueDate As Boolean = False
        Dim bolNewRevisionRcvdDate As Boolean = False
        Dim tmpdate As Date
        Dim oTecDoc As New MUSTER.BusinessLogic.pTecDoc
        Dim oTecAct As New MUSTER.BusinessLogic.pTecAct
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim inspCitation As New MUSTER.BusinessLogic.pInspectionCitation
        Dim CantNFAReason As String

        Try
            If cmbDocument.Text = "" Then
                MsgBox("Please select an Document")
                Exit Sub
            End If

            If Date.Compare(oLustDocument.DueDate, CDate("01/01/0001")) <> 0 Then
                If Date.Compare(oLustDocument.DueDate, oLustDocument.IssueDate) < 0 Then
                    MsgBox("Due Date Cannot Be Before Issue Date")
                    Exit Sub
                End If
            End If
            If Date.Compare(oLustDocument.DocClosedDate, CDate("01/01/0001")) <> 0 Then
                If Date.Compare(oLustDocument.DocClosedDate, oLustDocument.IssueDate) < 0 Then
                    MsgBox("Closed Date Cannot Be Before Issue Date")
                    Exit Sub
                End If
            End If
            If Date.Compare(oLustDocument.DocFinancialDate, CDate("01/01/0001")) <> 0 Then
                If Date.Compare(oLustDocument.DocFinancialDate, oLustDocument.IssueDate) < 0 Then
                    MsgBox("Sent To Financial Date Cannot Be Before Issue Date")
                    Exit Sub
                End If
            End If
            If Date.Compare(oLustDocument.DocRcvDate, CDate("01/01/0001")) <> 0 Then
                If Date.Compare(oLustDocument.DocRcvDate, oLustDocument.IssueDate) < 0 Then
                    MsgBox("Received Date Cannot Be Before Issue Date")
                    Exit Sub
                End If
            End If
            If oLustDocument.DocRcvDate <> "01/01/0001" And oLustDocument.DocClosedDate <> "01/01/0001" Then
                If oLustDocument.DocClosedDate < oLustDocument.DocRcvDate Then
                    MsgBox("Closed Date Cannot Be Before Received Date")
                    Exit Sub
                End If
            End If
            bolCantNFA = False

            oTecDoc.Retrieve(cmbDocument.SelectedValue)
            If oLustDocument.IsDirty Then
                If oLustDocument.ID <= 0 Then
                    bolNewDocument = True
                    If oLustDocument.DueDate = "01/01/0001" Then
                        If oTecDoc.DocType <> 917 Then
                            oLustDocument.DueDate = DateAdd(DateInterval.Day, 45, oLustDocument.IssueDate)
                        End If
                    End If
                    If InStr(oTecDoc.Name, "NFA", CompareMethod.Text) > 0 Then
                        'Disallow the NFA Document creation if ANY of the following conditions are true:
                        CantNFAReason = "Can't NFA Due To The Following:                      " & vbCrLf
                        CantNFAReason &= "     " & vbCrLf

                        CantNFADocumentID = oTecDoc.GetCantNFADocID

                        '    1.	The tank owner owes tank fees

                        If oFeeInvoice.GetCurrentBalance_Facility(oLustEvent.FacilityID) > 0 Then
                            'Cant NFA
                            bolCantNFA = True
                            CantNFAReason &= "     - Facility Owes Fees " & vbCrLf
                            If oTecDoc.Name.Trim() = "NFA" Then
                                contNFA = True
                            End If
                        End If

                        '    2.	Compliance and Enforcement has any outstanding violations due to a site inspection
                        '    4.	The owner has not complied with an Agreed Order
                        If inspCitation.CheckCitationExists(CDate("01/01/0001"), oLustEvent.FacilityID, False) Then
                            bolCantNFA = True
                            CantNFAReason &= "     - Owner Has C&E Issues " & vbCrLf
                        End If

                        '    5.	All invoices have not been submitted and processed
                        If Me.oLustEvent.OpenInvoices(oLustDocument.EventId) > 0 Then
                            bolCantNFA = True
                            CantNFAReason &= "     - Open Financial Invoices " & vbCrLf
                        End If

                        '    6.	All documents have not been submitted and processed
                        If Me.oLustEvent.OpenDocuments(oLustDocument.EventId) > 0 Then
                            bolCantNFA = True
                            CantNFAReason &= "     - Open Lust Documents " & vbCrLf
                        End If

                        '    3.	Tank closure is required and has not been completed
                        Dim ds As DataSet = oOwner.RunSQLQuery("SELECT COUNT(CLOSURE_ID) from tblCLO_CLOSURE WHERE CLOSURE_STATUS NOT IN (868, 869) AND DELETED = 0 AND FACILITY_ID = " + oLustEvent.FacilityID.ToString)
                        If ds.Tables(0).Rows(0)(0) > 0 Then
                            bolCantNFA = True
                            CantNFAReason &= "     - Owner has Closure Events which are not closed or cancelled " & vbCrLf
                        End If

                        If contNFA And CantNFADocumentID > 0 Then
                            If MsgBox("Facility owes fees. Do you want to continue to NFA?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                oLustDocument.DocClosedDate = Now.Date
                                CallingForm.Tag = 2
                            Else
                                If bolCantNFA = True Then
                                    CantNFAReason &= "     " & vbCrLf
                                    CantNFAReason &= "     " & vbCrLf

                                    oLustDocument.DocClass = CantNFADocumentID
                                    oLustDocument.DocumentType = CantNFADocumentID

                                    If CantNFADocumentID > 0 Then
                                        MsgBox(CantNFAReason & " Can't NFA Document Generated", MsgBoxStyle.OKOnly, "Can't NFA")
                                    Else
                                        MsgBox(CantNFAReason & " No Can't NFA Document Defined", MsgBoxStyle.OKOnly, "Can't NFA")
                                        Exit Sub
                                    End If
                                Else
                                    oLustDocument.DocClosedDate = Now.Date
                                    CallingForm.Tag = 2
                                End If
                            End If
                        End If
                    End If
                End If

                If oLustDocument.DocFinancialDate > tmpdate Then
                    bolNewSentToFinancial = oLustDocument.IsDirtySentToFinancial
                End If
                If oLustDocument.DocClosedDate > tmpdate Then
                    bolNewClosedDate = oLustDocument.IsDirtyClosedDate
                End If
                If oLustDocument.DocRcvDate > tmpdate Then
                    bolNewRecievedDate = oLustDocument.IsDirtyRecievedDate
                End If
                If oLustDocument.EXTENSIONDATE > tmpdate Then
                    bolNewExtensionDate = oLustDocument.IsDirtyExtensionDate
                End If
                If oLustDocument.DueDate > tmpdate Then
                    bolNewDueDate = oLustDocument.IsDirtyDueDate
                End If
                If oLustDocument.REV1EXTENSIONDATE > tmpdate Then
                    bolNewRevisionDueDate = oLustDocument.IsDirtyREV1Date
                End If
                If oLustDocument.REV1RECEIVEDDATE > tmpdate Then
                    bolNewRevisionRcvdDate = oLustDocument.IsDirtyREV1RecvdDate
                End If
                If oLustDocument.REV2EXTENSIONDATE > tmpdate Then
                    bolNewRevisionDueDate = oLustDocument.IsDirtyREV2Date
                End If
                If oLustDocument.REV2RECEIVEDDATE > tmpdate Then
                    bolNewRevisionRcvdDate = oLustDocument.IsDirtyREV2RecvdDate
                End If

                If oLustDocument.ID <= 0 Then
                    oLustDocument.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                End If


                oLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                ProcessDocumentsAndCalendarEntries(oLustDocument, bolNewDocument, bolNewSentToFinancial, bolNewClosedDate, bolNewRecievedDate, bolNewExtensionDate, bolNewDueDate, bolNewRevisionDueDate, bolNewRevisionRcvdDate, bolNewSentToFinancial)
            End If

            ProcessComments()

            Me.Close()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally

        End Try
    End Sub

    Private Sub ProcessDocumentsAndCalendarEntries(ByVal oLocalLustDocument As MUSTER.BusinessLogic.pLustEventDocument, ByVal bolNewDoc As Boolean, ByVal bolNewFinancial As Boolean, ByVal bolNewClosed As Boolean, ByVal bolNewRecieved As Boolean, ByVal bolNewExtension As Boolean, ByVal bolNewDue As Boolean, ByVal bolNewRevisionDue As Boolean, ByVal bolNewRevisionRvcd As Boolean, Optional ByVal createCalEntry As Boolean = False)
        Dim nDocTrigger As Int64
        Dim nDocType As Integer
        Dim tmpDate As Date
        Dim bolCreateDoc As Boolean = False
        Dim oLocalLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
        Dim oLustRemediation As New MUSTER.BusinessLogic.pLustRemediation
        Dim oDocument As New MUSTER.BusinessLogic.pTecDoc
        Dim otmpLustDoc As New MUSTER.BusinessLogic.pLustEventDocument
        Try
            oDocument.Retrieve(oLocalLustDocument.DocClass)
            nDocTrigger = oDocument.Trigger_Field
            nDocType = oDocument.DocType

            If bolNewDoc Then 'Creating a new Document
                If nDocType = 917 Then ' letter
                    Select Case nDocTrigger 'Generate Letter
                        Case 1134   'Due Date
                            If oLocalLustDocument.DueDate <> tmpDate Then
                                bolCreateDoc = True
                            End If
                        Case 1135   'Extension Date
                            If oLocalLustDocument.EXTENSIONDATE <> tmpDate Then
                                bolCreateDoc = True
                            End If
                        Case 1136   'Received Date
                            If oLocalLustDocument.DocRcvDate <> tmpDate Then
                                bolCreateDoc = True
                            End If
                        Case 1137   'To Financial
                            If oLocalLustDocument.DocFinancialDate <> tmpDate Then
                                bolCreateDoc = True
                            End If
                        Case 1138   'Closed Date
                            If oLocalLustDocument.DocClosedDate <> tmpDate Then
                                bolCreateDoc = True
                            End If
                        Case Else
                            bolCreateDoc = True
                    End Select
                Else
                    bolCreateDoc = True
                End If
            Else
                If nDocType = 917 Then ' letter
                    Select Case nDocTrigger 'Generate Letter
                        Case 1134   'Due Date
                            If bolNewDue Then
                                bolCreateDoc = True
                            End If
                        Case 1135   'Extension Date
                            If bolNewExtension Then
                                bolCreateDoc = True
                            End If
                        Case 1136   'Received Date
                            If bolNewRecieved Then
                                bolCreateDoc = True
                            End If
                        Case 1137   'To Financial
                            If bolNewFinancial Then
                                bolCreateDoc = True
                            End If
                        Case 1138   'Closed Date
                            If bolNewClosed Then
                                bolCreateDoc = True
                            End If
                    End Select
                End If
            End If

            If bolCreateDoc Then
                If nDocType = 917 Then ' letter
                    oLocalLustDocument.DocumentID = CreateDocument(oLocalLustDocument, bolNewDoc, nDocType)

                    'Set Tag in parent form to let it know a document was generated.  If a Filename exists.
                    If Trim(oDocument.FileName) > "" Then
                        If IsNumeric(CallingForm.Tag) Then
                            If CallingForm.Tag <> 2 Then
                                CallingForm.Tag = 1
                            End If
                        Else
                            CallingForm.Tag = 1
                        End If
                    End If

                End If
                'ProcessCalendarEntries(nDocType, bolNewFinancial, bolNewClosed, bolNewRecieved, bolNewDue, bolNewRevisionDue, bolNewRevisionRvcd)
                'Else
                'If nDocType = 917 Then
                'ProcessCalendarEntries(nDocType, bolNewFinancial, bolNewClosed, bolNewRecieved, bolNewDue, bolNewRevisionDue, bolNewRevisionRvcd)
                'End If
            End If

            ProcessCalendarEntries(oDocument, oLocalLustDocument, nDocType, bolNewFinancial, bolNewClosed, bolNewRecieved, bolNewDue, bolNewRevisionDue, bolNewRevisionRvcd, bolNewExtension, createCalEntry)


            'Handle the Aftermath.....


            If bolNewFinancial Then
                If (oLocalLustDocument.DocClass = 804) Then   'TK SOW/CE
                    ProcessTKSOW(oLocalLustDocument.EventId)
                Else
                    If (oLocalLustDocument.DocClass = 1611) Then 'TK SOW/CE Temporary
                        ProcessTKSOWTEMP(oLocalLustDocument.EventId)
                    Else
                        If (oLustEvent.MGPTFStatus = 618 Or oLustEvent.MGPTFStatus = 619 Or oLustEvent.MGPTFStatus = 621) Then
                            GenerateAdditionalDocs(oLocalLustDocument.EventId, oLocalLustDocument.AssocActivity, oLocalLustDocument.DocClass, bolNewFinancial, createCalEntry)
                        End If
                    End If
                End If
            End If


                If bolNewClosed Then
                    If (oLustEvent.MGPTFStatus = 617 Or oLustEvent.MGPTFStatus = 620) Then
                        ' Generate Additional Docs
                        GenerateAdditionalDocs(oLocalLustDocument.EventId, oLocalLustDocument.AssocActivity, oLocalLustDocument.DocClass)
                    End If

                End If


                If bolNewClosed And oLocalLustDocument.DocRcvDate <> tmpDate Then
                    'If bolNewRecieved Then
                    'if the Document is GWS Rpt X of Y, and the Received date is greater than the 
                    'LUST Event’s Last GWS, replace the LUST Events Last GWS with the Received 
                    'Date 
                    'If oLocalLustDocument.DocClass = 729 Then
                    If oDocument.Name.ToUpper.StartsWith("GWS RPT") Then
                        oLustEvent.LastGWS = oLocalLustDocument.DocRcvDate
                    End If

                    'if the Document is LDRs, and the Received date is greater than the LUST 
                    'Event’s Last LDR, replace the LUST Events Last LDR with the Received Date 
                    If oLocalLustDocument.DocClass = 734 Then
                        oLustEvent.LastLDR = oLocalLustDocument.DocRcvDate
                        'oLustEvent.Save()
                    End If

                    ' if the Document is PTT, and the Received date is greater than the LUST 
                    'Event’s Last PTT, replace the LUST Events Last PTT with the Received Date 
                    If oLocalLustDocument.DocClass = 782 Then
                        oLustEvent.LastPTT = oLocalLustDocument.DocRcvDate
                        'oLustEvent.Save()
                    End If
                End If


                If bolNewRecieved Then
                    If oLocalLustDocument.DocClass = 798 Then ' System Startup Date
                        'If the Received Date has been entered or modified and the Document is 
                        'System Startup Date, enter the Received Date in the System Start Date 
                        '(for the Remediation System) AND enter the Received Date in the Start 
                        'Date for the associated REM Activity
                        oLocalLustActivity.Retrieve(oLocalLustDocument.AssocActivity)
                        oLustRemediation.Retrieve(oLocalLustActivity.RemSystemID)
                        oLocalLustActivity.Started = oLocalLustDocument.DocRcvDate
                        oLustRemediation.DateInUse = oLocalLustDocument.DocRcvDate
                        oLocalLustActivity.ModifiedBy = MusterContainer.AppUser.ID
                        oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        oLustRemediation.ModifiedBy = MusterContainer.AppUser.ID
                        oLustRemediation.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                        'If the Received Date has been entered or modified and the Document is System 
                        'Startup Date, enter the Received Date + 30 days in the Due Date of the System 
                        'Startup Rpt Document.
                        otmpLustDoc.Retrieve(oLocalLustActivity.ActivityID, 799)
                        If otmpLustDoc.ID > 0 Then
                            otmpLustDoc.DueDate = DateAdd(DateInterval.Day, 30, oLocalLustDocument.DocRcvDate)
                            otmpLustDoc.ModifiedBy = MusterContainer.AppUser.ID
                            otmpLustDoc.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If
                            'Create Due-To-Me calendar entry for System Startup Rpt Document.
                            Dim strTaskDesc As String = String.Empty
                            Dim nColorCode As Integer
                            strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & "System Startup Rpt" & " - New Due Date"
                            otmpLustDoc.MarkDueToMeCompleted_ByDesc(otmpLustDoc.ID, strTaskDesc)
                            Dim oUser As New MUSTER.BusinessLogic.pUser
                            oUser.Retrieve(oLustEvent.PM)

                            Dim pCal As New MUSTER.BusinessLogic.pCalendar
                            Dim oCalInfo As MUSTER.Info.CalendarInfo
                            oCalInfo = New MUSTER.Info.CalendarInfo(0, _
                                    Now(), _
                                    otmpLustDoc.DueDate, _
                                    nColorCode, _
                                    strTaskDesc, _
                                    oUser.ID, _
                                    "SYSTEM", _
                                    "", _
                                    True, _
                                    False, _
                                    False, _
                                    False, _
                                    "SYSTEM", _
                                    Now(), _
                                    "SYSTEM", _
                                    Now())

                            oCalInfo.OwningEntityID = otmpLustDoc.ID
                            oCalInfo.OwningEntityType = UIUtilsGen.EntityTypes.LustDocument
                            oCalInfo.IsDirty = True
                            pCal.Add(oCalInfo)
                            pCal.Save()
                        End If
                    End If

                    If oLocalLustDocument.DocClass = 797 Then ' System Shutdown By
                        'If the Received Date has been entered or modified and the Document is 
                        'System Shutdown By, enter the Received Date in the Technically Closed 
                        'Date for the associated REM acitivity 
                        oLocalLustActivity.Retrieve(oLocalLustDocument.AssocActivity)
                        oLocalLustActivity.Completed = oLocalLustDocument.DocRcvDate
                    oLocalLustActivity.ModifiedBy = MusterContainer.AppUser.ID
                    'Added by Hua Cao 07/03/2008  
                    'Issue # 3179
                    'oLustEvent.FacilityID = 9973
                    'oLocalLustActivity.EventID = 1439, AssocActivity = 2472
                    'select * from dbo.tblTEC_EVENT_ACTIVITY_DOCUMENT where event_id = '1439' 
                    '       and event_activity_id = '2472'
                    ' Loop to see if there is any doc has Date_Closed is null, if yes, mark hasOpenDoc to True
                    'If hasOpenDoc = True
                    'oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal, hasOpenDoc)
                    ' Else do the following line
                    If oLocalLustActivity.GetOpenDocumentCount(oLocalLustActivity.ActivityID) > 0 Then
                        oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal, True)
                    Else

                        'If oLocalLustActivity.GetNoClosedDateDoc(oLocalLustActivity.EventID, oLocalLustActivity.ActivityID) > 0 Then
                        'oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal, True)
                        'Else
                        oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                    End If
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If

                End If

                If oLocalLustDocument.IsDirty Then
                    oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                    oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try


    End Sub

    Private Sub ProcessTKSOW(ByVal nEventId As Long)
        'c.	If the Sent To Financial Date has been entered and was previously 
        ' empty and the Document is TK SOW/CE, creating a REM-DUAL PHASE 
        ' Activity (with no Start Date) and creating the following Documents 
        ' for the REM-DUAL Phase Activity:  (692)
        '   i. 	    System Startup Date (798)
        '   ii.	    System Startup Rpt  (799)
        '   iii.	YR1 Tri Rpt 1       (818)
        '   iv.	    YR1 Tri Rpt 2       (819)
        '   v.	    YR2 Cont SOW/CE     (821)
        Dim oLocalLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
        Dim oLocalLustDocument As MUSTER.BusinessLogic.pLustEventDocument
        Dim bolRemSystemDone As Boolean = False
        Dim bolInsertDocs As Boolean = False

        Try

            oLocalLustActivity.Retrieve(0)
            oLocalLustActivity.EventID = nEventId
            oLocalLustActivity.FacilityID = oLustActivity.FacilityID
            oLocalLustActivity.RemSystemID = 0
            oLocalLustActivity.Type = 692   'REM-DUAL Phase Activity
            oLocalLustActivity.ModifiedBy = MusterContainer.AppUser.ID
            oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            '========
            While bolRemSystemDone = False
                AddRemediationSystem(oLocalLustActivity.ActivityID)
                oLocalLustActivity.AgeThreshold = 0
                oLocalLustActivity.Retrieve(oLocalLustActivity.ActivityID)
                If oLocalLustActivity.RemSystemID > 0 Then
                    bolRemSystemDone = True
                    bolInsertDocs = True
                Else
                    If MsgBox("No remediation system was assigned.  Do you want to do this now?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        bolRemSystemDone = False
                    Else
                        'oLocalLustActivity.Deleted = True
                        oLocalLustActivity.CreatedBy = MusterContainer.AppUser.ID
                        oLocalLustActivity.ModifiedBy = MusterContainer.AppUser.ID
                        oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                        bolInsertDocs = True
                        bolRemSystemDone = True
                        End If
                End If

            End While
            '========

            '    If bolInsertDocs Then
            oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
            oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
            oLocalLustDocument.DocClass = 798
            oLocalLustDocument.DocumentType = 798
            oLocalLustDocument.EventId = nEventId
            oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
            oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
            oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

            oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
            oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
            oLocalLustDocument.DocClass = 799
            oLocalLustDocument.DocumentType = 799
            oLocalLustDocument.EventId = nEventId
            oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
            oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
            oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

            oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
            oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
            oLocalLustDocument.DocClass = 818
            oLocalLustDocument.DocumentType = 818
            oLocalLustDocument.EventId = nEventId
            oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
            oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
            oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

            oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
            oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
            oLocalLustDocument.DocClass = 819
            oLocalLustDocument.DocumentType = 819
            oLocalLustDocument.EventId = nEventId
            oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
            oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
            oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

            oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
            oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
            oLocalLustDocument.DocClass = 821
            oLocalLustDocument.DocumentType = 821
            oLocalLustDocument.EventId = nEventId
            oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
            oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
            oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

            '    End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ProcessTKSOWTEMP(ByVal nEventId As Long)
        'c.	If the Sent To Financial Date has been entered and was previously 
        ' empty and the Document is TK SOW/CE TEMPORARY, creating a REM-DUAL PHASE TEMPORARY
        ' Activity (with no Start Date) and creating the following Documents 
        ' for the REM-DUAL Phase Activity:  (1561)
        '   i. 	    System Startup Date (798)
        '   ii.	    System Startup Rpt  (799)
        '   iii.	Temp System Rpt 1 of 6       (1573)
        '   iv.	    Temp System Rpt 2 of 6       (1574)
        '   v.	    Temp System Rpt 3 of 6       (1575)
        '   vi.     Temp System Rpt 4 of 6       (1576)
        '   vii.    Temp System Rpt 5 of 6       (1577)
        '   vii.    Temp System Rpt 6 of 6       (1578)
        Dim oLocalLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
        Dim oLocalLustDocument As MUSTER.BusinessLogic.pLustEventDocument
        Dim bolRemSystemDone As Boolean = False
        Dim bolInsertDocs As Boolean = False

        Try

            oLocalLustActivity.Retrieve(0)
            oLocalLustActivity.EventID = nEventId
            oLocalLustActivity.FacilityID = oLustActivity.FacilityID
            oLocalLustActivity.RemSystemID = 0
            oLocalLustActivity.Type = 1561   'REM-DUAL Phase Temporary Activity
            oLocalLustActivity.ModifiedBy = MusterContainer.AppUser.ID
            oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            '========
            While bolRemSystemDone = False
                AddRemediationSystem(oLocalLustActivity.ActivityID)
                oLocalLustActivity.AgeThreshold = 0
                oLocalLustActivity.Retrieve(oLocalLustActivity.ActivityID)
                If oLocalLustActivity.RemSystemID > 0 Then
                    bolRemSystemDone = True
                    bolInsertDocs = True
                Else
                    If MsgBox("No remediation system was assigned.  Do you want to do this now?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        bolRemSystemDone = False
                    Else
                        oLocalLustActivity.Deleted = True
                        oLocalLustActivity.ModifiedBy = MusterContainer.AppUser.ID
                        oLocalLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                        bolRemSystemDone = True
                    End If
                End If

            End While
            '========

            If bolInsertDocs Then
                oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
                oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
                oLocalLustDocument.DocClass = 798
                oLocalLustDocument.DocumentType = 798
                oLocalLustDocument.EventId = nEventId
                oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
                oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

                oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
                oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
                oLocalLustDocument.DocClass = 799
                oLocalLustDocument.DocumentType = 799
                oLocalLustDocument.EventId = nEventId
                oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
                oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

                oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
                oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
                oLocalLustDocument.DocClass = 1573
                oLocalLustDocument.DocumentType = 1573
                oLocalLustDocument.EventId = nEventId
                oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
                oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

                oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
                oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
                oLocalLustDocument.DocClass = 1574
                oLocalLustDocument.DocumentType = 1574
                oLocalLustDocument.EventId = nEventId
                oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
                oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

                oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
                oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
                oLocalLustDocument.DocClass = 1575
                oLocalLustDocument.DocumentType = 1575
                oLocalLustDocument.EventId = nEventId
                oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
                oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

                oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
                oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
                oLocalLustDocument.DocClass = 1576
                oLocalLustDocument.DocumentType = 1576
                oLocalLustDocument.EventId = nEventId
                oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
                oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

                oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
                oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
                oLocalLustDocument.DocClass = 1577
                oLocalLustDocument.DocumentType = 1577
                oLocalLustDocument.EventId = nEventId
                oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
                oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

                oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
                oLocalLustDocument.AssocActivity = oLocalLustActivity.ActivityID
                oLocalLustDocument.DocClass = 1578
                oLocalLustDocument.DocumentType = 1578
                oLocalLustDocument.EventId = nEventId
                oLocalLustDocument.FacilityID = oLocalLustActivity.FacilityID
                oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)

            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub AddRemediationSystem(ByVal nActivityID As Int64)
        Dim frmRemSysList As New RemediationSystemList
        Try

            frmRemSysList.CallingForm = Me
            frmRemSysList.Mode = 0 ' Add
            frmRemSysList.EventActivityID = nActivityID

            frmRemSysList.ShowDialog()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Add Remediation System" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmRemSysList = Nothing
        End Try
    End Sub

    Private Sub GenerateAdditionalDocs(ByVal nEventId As Int64, ByVal nAssocActivity As Int64, ByVal nDocClass As Int64, Optional ByVal bolFinancial As Boolean = False, Optional ByVal createCalEntry As Boolean = False)
        Dim oTecDocument As New MUSTER.BusinessLogic.pTecDoc

        Try

            oTecDocument.Retrieve(nDocClass)
            If oTecDocument.Auto_Doc_1 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_1, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_2 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_2, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_3 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_3, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_4 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_4, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_5 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_5, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_6 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_6, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_7 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_7, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_8 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_8, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_9 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_9, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If
            If oTecDocument.Auto_Doc_10 > 0 Then
                ProcessAdditionalDocs(oTecDocument.Auto_Doc_10, nEventId, nAssocActivity, bolFinancial, createCalEntry)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ProcessAdditionalDocs(ByVal DocID As Int64, ByVal nEventId As Int64, ByVal nAssocActivity As Int64, Optional ByVal bolFinancial As Boolean = False, Optional ByVal createCalEntry As Boolean = False)
        Dim oLocalLustDocument As MUSTER.BusinessLogic.pLustEventDocument
        oLocalLustDocument = New MUSTER.BusinessLogic.pLustEventDocument
        oLocalLustDocument.AssocActivity = nAssocActivity
        oLocalLustDocument.DocClass = DocID
        oLocalLustDocument.DocumentType = DocID
        oLocalLustDocument.EventId = nEventId
        oLocalLustDocument.FacilityID = nAssocActivity
        oLocalLustDocument.ModifiedBy = MusterContainer.AppUser.ID
        If bolFinancial Then
            oLocalLustDocument.IssueDate = Date.Now
        End If
        oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
        If Not UIUtilsGen.HasRights(returnVal) Then
            Exit Sub
        End If
        'ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False)
        ProcessDocumentsAndCalendarEntries(oLocalLustDocument, True, False, False, False, False, False, False, False, createCalEntry)
    End Sub

    Private Function CreateDocument(ByVal oLocalLustDocument As MUSTER.BusinessLogic.pLustEventDocument, ByVal bolNewDoc As Boolean, ByVal DocType As Integer) As Long
        Dim oLetter As New Reg_Letters
        Dim strShortName As String
        Dim strLongName As String
        Dim strTemplate As String
        Dim oDocument As New MUSTER.BusinessLogic.pTecDoc

        Try
            oDocument.Retrieve(oLocalLustDocument.DocClass)
            strTemplate = oDocument.FileName

            If strTemplate <> "" And strTemplate <> "Default" Then
                strLongName = oDocument.Name
                strShortName = strLongName
                strShortName = strShortName.Replace(" ", "")
                strShortName = strShortName.Replace("a", "")
                strShortName = strShortName.Replace("e", "")
                strShortName = strShortName.Replace("i", "")
                strShortName = strShortName.Replace("o", "")
                strShortName = strShortName.Replace("u", "")
                strShortName = strShortName.Replace("/", "")
                strShortName = strShortName.Replace("\", "")
                strShortName = strShortName.Replace(".", "")
                strShortName = strShortName.Replace("'", "")

                oLetter.GenerateTechLetter(oLustEvent.FacilityID, strLongName, Mid(strShortName, 1, 8), strLongName, strTemplate, oLocalLustDocument.DueDate, oLustEvent.ID, oOwner, 0, oLustEvent.EVENTSEQUENCE, UIUtilsGen.EntityTypes.LUST_Event)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try


    End Function


    Private Function ProcessCalendarEntries(ByVal oDocument As MUSTER.BusinessLogic.pTecDoc, ByVal oLocalLustDocument As MUSTER.BusinessLogic.pLustEventDocument, ByVal DocType As Integer, ByVal bolNewFinancial As Boolean, ByVal bolNewClosed As Boolean, ByVal bolNewRecieved As Boolean, ByVal bolNewDue As Boolean, ByVal bolNewRevisionDue As Boolean, ByVal bolNewRevisionRvcd As Boolean, ByVal bolNewExtension As Boolean, Optional ByVal createCalEntry As Boolean = False)
        Dim ocalendar1 As New MUSTER.BusinessLogic.pCalendar
        Dim ocalendar2 As New MUSTER.BusinessLogic.pCalendar
        Dim ocalendar3 As New MUSTER.BusinessLogic.pCalendar
        Dim ocalendar4 As New MUSTER.BusinessLogic.pCalendar
        Dim ocalendar5 As New MUSTER.BusinessLogic.pCalendar
        Dim ocalendar6 As New MUSTER.BusinessLogic.pCalendar
        Dim ocalendar7 As New MUSTER.BusinessLogic.pCalendar
        Dim ocalendar8 As New MUSTER.BusinessLogic.pCalendar
        Dim ocalendar As New MUSTER.BusinessLogic.pCalendar
        Dim oUser As New MUSTER.BusinessLogic.pUser

        Dim tmpDate As Date
        Dim bolAddCalendarEntry As Boolean = False
        Dim bolCalendarEntryAdded As Boolean = False
        Dim dtNotificationDate As Date = Now()
        Dim dtDueDate As Date
        Dim nColorCode
        Dim strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name 'cmbDocument.Text
        Dim strUserID As String = ""
        Dim strSourceUserID As String = "SYSTEM"
        Dim strGroupID As String = ""
        Dim bolDuetoMe As Boolean = False
        Dim bolToDo As Boolean = False
        Dim bolCompleted As Boolean = False
        Dim bolDeleted As Boolean = False
        Dim oCalendarInfo1 As MUSTER.Info.CalendarInfo
        Dim oCalendarInfo2 As MUSTER.Info.CalendarInfo
        Dim oCalendarInfo3 As MUSTER.Info.CalendarInfo
        Dim oCalendarInfo4 As MUSTER.Info.CalendarInfo
        Dim oCalendarInfo5 As MUSTER.Info.CalendarInfo
        Dim oCalendarInfo6 As MUSTER.Info.CalendarInfo
        Dim oCalendarInfo7 As MUSTER.Info.CalendarInfo
        Dim oCalendarInfo8 As MUSTER.Info.CalendarInfo
        Dim oCalendarInfo As MUSTER.Info.CalendarInfo

        Try
            oUser.Retrieve(oLustEvent.PM)

            'ocalendar.Retrieve(oLustDocument.EntityID, oLustDocument.ID, Nothing, Nothing)

            If bolNewDue And oLustDocument.DueDate > tmpDate Then
                '•	Create a Due To Me Calendar entry for the user on the Due date of the Document, unless it is a Task, 
                strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name & " - New Due Date"
                oLustDocument.MarkDueToMeCompleted_ByDesc(oLocalLustDocument.ID, strTaskDesc)
                If DocType = 919 Then ' task
                    bolToDo = True
                    bolDuetoMe = False
                Else
                    bolDuetoMe = True
                    bolToDo = False
                End If

                dtDueDate = oLustDocument.DueDate
                strUserID = oUser.ID

                oCalendarInfo3 = New MUSTER.Info.CalendarInfo(0, _
                        dtNotificationDate, _
                        dtDueDate, _
                        nColorCode, _
                        strTaskDesc, _
                        strUserID, _
                        strSourceUserID, _
                        strGroupID, _
                        bolDuetoMe, _
                        bolToDo, _
                        bolCompleted, _
                        bolDeleted, _
                        "SYSTEM", _
                        Now(), _
                        "SYSTEM", _
                        Now())

                oCalendarInfo3.OwningEntityID = oLocalLustDocument.ID
                oCalendarInfo3.OwningEntityType = UIUtilsGen.EntityTypes.LustDocument
                oCalendarInfo3.IsDirty = True
                ocalendar3.Add(oCalendarInfo3)
                ocalendar3.Flush()

            End If

            If bolNewRevisionDue And (oLustDocument.REV1EXTENSIONDATE > tmpDate Or oLustDocument.REV2EXTENSIONDATE > tmpDate) Then
                '•	Remove any existing associated To Do Calendar entries
                oLustDocument.MarkToDoCompleted(oLocalLustDocument.ID)

                '•	Create Due To Me Calendar entry for the user on the Revision Due Date
                bolToDo = False
                bolDuetoMe = True
                strUserID = oUser.ID
                strGroupID = ""
                If oLustDocument.REV2EXTENSIONDATE > tmpDate Then
                    dtDueDate = oLustDocument.REV2EXTENSIONDATE
                Else
                    dtDueDate = oLustDocument.REV1EXTENSIONDATE
                End If

                strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name & " - Revision Due"

                oCalendarInfo1 = New MUSTER.Info.CalendarInfo(0, _
                                                dtNotificationDate, _
                                                dtDueDate, _
                                                nColorCode, _
                                                strTaskDesc, _
                                                strUserID, _
                                                strSourceUserID, _
                                                strGroupID, _
                                                bolDuetoMe, _
                                                bolToDo, _
                                                bolCompleted, _
                                                bolDeleted, _
                                                "SYSTEM", _
                                                Now(), _
                                                "SYSTEM", _
                                                Now())

                oCalendarInfo1.OwningEntityID = oLocalLustDocument.ID
                oCalendarInfo1.OwningEntityType = UIUtilsGen.EntityTypes.LustDocument
                oCalendarInfo1.IsDirty = True
                ocalendar1.Add(oCalendarInfo1)
                ocalendar1.Flush()

            End If


            If bolNewRevisionRvcd And (oLustDocument.REV1RECEIVEDDATE > tmpDate Or oLustDocument.REV2RECEIVEDDATE > tmpDate) Then
                '•	Remove any existing associated Due To Me Calendar entries
                oLustDocument.MarkDueToMeCompleted(oLocalLustDocument.ID)

                '•	Create Due To Me Calendar entry for the user on the Revision Due Date
                bolToDo = True
                bolDuetoMe = False
                strUserID = oUser.ID
                strGroupID = ""
                If oLustDocument.REV2RECEIVEDDATE > tmpDate Then
                    dtDueDate = oLustDocument.REV2RECEIVEDDATE
                Else
                    dtDueDate = oLustDocument.REV1RECEIVEDDATE
                End If

                strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name & " - Revision Received"

                oCalendarInfo2 = New MUSTER.Info.CalendarInfo(0, _
                        dtNotificationDate, _
                        dtDueDate, _
                        nColorCode, _
                        strTaskDesc, _
                        strUserID, _
                        strSourceUserID, _
                        strGroupID, _
                        bolDuetoMe, _
                        bolToDo, _
                        bolCompleted, _
                        bolDeleted, _
                        "SYSTEM", _
                        Now(), _
                        "SYSTEM", _
                        Now())

                oCalendarInfo2.OwningEntityID = oLocalLustDocument.ID
                oCalendarInfo2.OwningEntityType = UIUtilsGen.EntityTypes.LustDocument
                oCalendarInfo2.IsDirty = True
                ocalendar2.Add(oCalendarInfo2)
                ocalendar2.Flush()

            End If

            If bolNewExtension And oLustDocument.EXTENSIONDATE > tmpDate Then
                '•	Create a Due To Me Calendar entry for the user on the Due date of the Document, unless it is a Task, 
                strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name & " - New Due Date"
                oLustDocument.MarkDueToMeCompleted_ByDesc(oLocalLustDocument.ID, strTaskDesc)
                strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name & " - New Extension Date"
                If DocType = 919 Then ' task
                    bolToDo = True
                    bolDuetoMe = True
                Else
                    bolDuetoMe = True
                    bolToDo = False
                End If
                dtDueDate = oLustDocument.EXTENSIONDATE
                strUserID = oUser.ID

                ' #2895
                ' check if cal entry exists. if exists, modify cal due date else create new cal entry
                Dim oCalCol6 As MUSTER.Info.CalendarCollection = ocalendar6.RetrieveByOtherID(UIUtilsGen.EntityTypes.LustDocument, oLocalLustDocument.ID, strTaskDesc, "DESCRIPTION")
                If oCalCol6.Count > 0 Then
                    oCalendarInfo6 = oCalCol6.Item(oCalCol6.GetKeys(0))
                    oCalendarInfo6.DateDue = dtDueDate
                    oCalendarInfo6.UserId = strUserID
                Else
                    oCalendarInfo6 = New MUSTER.Info.CalendarInfo(0, _
                            dtNotificationDate, _
                            dtDueDate, _
                            nColorCode, _
                            strTaskDesc, _
                            strUserID, _
                            strSourceUserID, _
                            strGroupID, _
                            bolDuetoMe, _
                            bolToDo, _
                            bolCompleted, _
                            bolDeleted, _
                            "SYSTEM", _
                            Now(), _
                            "SYSTEM", _
                            Now())
                    oCalendarInfo6.OwningEntityID = oLocalLustDocument.ID
                    oCalendarInfo6.OwningEntityType = UIUtilsGen.EntityTypes.LustDocument
                End If
                oCalendarInfo6.IsDirty = True
                ocalendar6.Add(oCalendarInfo6)
                ocalendar6.Flush()
            End If
            If bolNewRecieved And oLustDocument.DocRcvDate > tmpDate Then
                '•	Remove any existing associated To Do or Due to Me Calendar entries
                oLustDocument.MarkToDoCompleted(oLocalLustDocument.ID)
                oLustDocument.MarkDueToMeCompleted(oLocalLustDocument.ID)
                strUserID = oUser.ID
                '•	Create a To Do Calendar entry for the user on the Received date of the Document
                bolToDo = True
                bolDuetoMe = False
                dtDueDate = oLustDocument.DocRcvDate
                strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name & " - Received"

                oCalendarInfo4 = New MUSTER.Info.CalendarInfo(0, _
                        dtNotificationDate, _
                        dtDueDate, _
                        nColorCode, _
                        strTaskDesc, _
                        strUserID, _
                        strSourceUserID, _
                        strGroupID, _
                        bolDuetoMe, _
                        bolToDo, _
                        bolCompleted, _
                        bolDeleted, _
                        "SYSTEM", _
                        Now(), _
                        "SYSTEM", _
                        Now())

                oCalendarInfo4.OwningEntityID = oLocalLustDocument.ID
                oCalendarInfo4.OwningEntityType = UIUtilsGen.EntityTypes.LustDocument
                oCalendarInfo4.IsDirty = True
                ocalendar4.Add(oCalendarInfo4)
                ocalendar4.Flush()

            End If

            If (bolNewFinancial Or createCalEntry) And oLustDocument.DocFinancialDate > tmpDate Then
                '•	Remove any existing associated To Do or Due to Me Calendar entries
                oLustDocument.MarkToDoCompleted(oLocalLustDocument.ID)
                oLustDocument.MarkDueToMeCompleted(oLocalLustDocument.ID)

                '•	Create To Do Calendar entry for the Financial Group on the Sent To Financial date
                bolToDo = True
                bolDuetoMe = False
                strUserID = ""
                strGroupID = "Financial"
                dtDueDate = oLustDocument.DocFinancialDate
                strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name & " - To Financial"

                oCalendarInfo5 = New MUSTER.Info.CalendarInfo(0, _
                                                dtNotificationDate, _
                                                dtDueDate, _
                                                nColorCode, _
                                                strTaskDesc, _
                                                strUserID, _
                                                strSourceUserID, _
                                                strGroupID, _
                                                bolDuetoMe, _
                                                bolToDo, _
                                                bolCompleted, _
                                                bolDeleted, _
                                                "SYSTEM", _
                                                Now(), _
                                                "SYSTEM", _
                                                Now())

                oCalendarInfo5.OwningEntityID = oLocalLustDocument.ID
                oCalendarInfo5.OwningEntityType = UIUtilsGen.EntityTypes.LustDocument
                oCalendarInfo5.IsDirty = True
                ocalendar5.Add(oCalendarInfo5)
                ocalendar5.Flush()

            End If


            If bolNewClosed And oLustDocument.DocClosedDate > tmpDate Then
                '•	Remove any existing associated To Do or Due to Me Calendar entries
                oLustDocument.MarkToDoCompleted(oLocalLustDocument.ID)
                oLustDocument.MarkDueToMeCompleted(oLocalLustDocument.ID)
            End If



            'If bolAddCalendarEntry = True Then

            '    'Create a Calendar Info object 

            '    oCalendarInfo = New MUSTER.Info.CalendarInfo(0, _
            '                                    dtNotificationDate, _
            '                                    dtDueDate, _
            '                                    nColorCode, _
            '                                    strTaskDesc, _
            '                                    strUserID, _
            '                                    strSourceUserID, _
            '                                    strGroupID, _
            '                                    bolDuetoMe, _
            '                                    bolToDo, _
            '                                    bolCompleted, _
            '                                    bolDeleted, _
            '                                    "sdfsdf", _
            '                                    Now(), _
            '                                    "asdf", _
            '                                    Now())

            '    oCalendarInfo.OwningEntityID = oLustDocument.ID
            '    oCalendarInfo.OwningEntityType = oLustDocument.EntityID
            '    oCalendarInfo.IsDirty = True
            '    ocalendar.Add(oCalendarInfo)
            '    ocalendar.Flush()
            'End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Function


    Private Sub SetDocumentType()
        Dim oDocument As New MUSTER.BusinessLogic.pTecDoc
        If IsNothing(cmbDocument.SelectedValue) Then Exit Sub
        oDocument.Retrieve(cmbDocument.SelectedValue)
        nCurrentDocType = oDocument.DocType

        dtDue.Enabled = True

        If (nCurrentDocType = 919 Or nCurrentDocType = 918) And dtDue.Checked = False And cmbDocument.Text.IndexOf("MDEQ SOW") <= -1 Then
            dtDue.Checked = True
            dtDue.Value = DateAdd(DateInterval.Day, 45, Now.Date)
        End If

        If nCurrentDocType = 919 Then
            dtReceived.Enabled = False
        Else
            dtReceived.Enabled = True
        End If

        If (nCurrentDocType = 919 Or nCurrentDocType = 918) Then
            dtToFinancial.Enabled = False
            dtExtension.Enabled = False
            gbRevision1.Enabled = False
            gbRevision2.Enabled = False
        Else
            If TFStatus = 617 Or TFStatus = 620 Then
                dtToFinancial.Enabled = False
            Else
                dtToFinancial.Enabled = True
            End If
            If Mode <> 0 Then
                dtExtension.Enabled = True
                gbRevision1.Enabled = True
                gbRevision2.Enabled = True
            End If
        End If
        If cmbDocument.Text.IndexOf("MDEQ SOW") > -1 Then
            dtDue.Enabled = False
            dtExtension.Enabled = False
            gbRevision1.Enabled = False
            gbRevision2.Enabled = False
            dtReceived.Enabled = False
        End If

    End Sub


    Private Sub Document_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        'If oLustDocument.IsDirty Then
        '    If MsgBox("Do you wish to save changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
        '        ProcessSaveEvent()
        '    End If
        'End If

    End Sub

    Private Sub Document_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If oLustDocument.IsDirty Then
            If MsgBox("Do you wish to save changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                ProcessSaveEvent()
            End If
        End If

    End Sub




End Class
