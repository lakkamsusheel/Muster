Public Class GenerateRefund
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Friend OwnerID As Int64

    Private oInvoice As New MUSTER.BusinessLogic.pFeeInvoice
    Private oInvoiceADLN As New MUSTER.Info.FeeInvoiceInfo
    Private oAddress As New MUSTER.BusinessLogic.pAddress
    Private bolLoading As Boolean
    Private oFeeBasis As New MUSTER.BusinessLogic.pFeeBasis

    Dim returnVal As String = String.Empty
    Dim TotalOverage As Single
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
    Friend WithEvents pnlGenerateRefundBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlGenRefundDetails As System.Windows.Forms.Panel
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblFromCheck As System.Windows.Forms.Label
    Friend WithEvents cmbFromCheck As System.Windows.Forms.ComboBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblTotalOverage As System.Windows.Forms.Label
    Friend WithEvents txtRefundAmount As System.Windows.Forms.TextBox
    Friend WithEvents lblTotalOverageValue As System.Windows.Forms.Label
    Friend WithEvents lblRefundAmount As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkOwner As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblZip As System.Windows.Forms.Label
    Friend WithEvents lblState As System.Windows.Forms.Label
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblAddress2 As System.Windows.Forms.Label
    Friend WithEvents lblAddress1 As System.Windows.Forms.Label
    Friend WithEvents cboZipCode As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboState As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboCity As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents txtAddress2 As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress1 As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlGenerateRefundBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlGenRefundDetails = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lblZip = New System.Windows.Forms.Label
        Me.lblState = New System.Windows.Forms.Label
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblAddress2 = New System.Windows.Forms.Label
        Me.lblAddress1 = New System.Windows.Forms.Label
        Me.cboZipCode = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboState = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboCity = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.txtAddress2 = New System.Windows.Forms.TextBox
        Me.txtAddress1 = New System.Windows.Forms.TextBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.chkOwner = New System.Windows.Forms.CheckBox
        Me.txtRefundAmount = New System.Windows.Forms.TextBox
        Me.lblRefundAmount = New System.Windows.Forms.Label
        Me.lblTotalOverageValue = New System.Windows.Forms.Label
        Me.cmbFromCheck = New System.Windows.Forms.ComboBox
        Me.lblFromCheck = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.lblComments = New System.Windows.Forms.Label
        Me.lblTotalOverage = New System.Windows.Forms.Label
        Me.pnlGenerateRefundBottom.SuspendLayout()
        Me.pnlGenRefundDetails.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.cboZipCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboState, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboCity, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlGenerateRefundBottom
        '
        Me.pnlGenerateRefundBottom.Controls.Add(Me.btnCancel)
        Me.pnlGenerateRefundBottom.Controls.Add(Me.btnSave)
        Me.pnlGenerateRefundBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlGenerateRefundBottom.Location = New System.Drawing.Point(0, 278)
        Me.pnlGenerateRefundBottom.Name = "pnlGenerateRefundBottom"
        Me.pnlGenerateRefundBottom.Size = New System.Drawing.Size(608, 40)
        Me.pnlGenerateRefundBottom.TabIndex = 4
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(307, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(227, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Save"
        '
        'pnlGenRefundDetails
        '
        Me.pnlGenRefundDetails.Controls.Add(Me.GroupBox1)
        Me.pnlGenRefundDetails.Controls.Add(Me.txtRefundAmount)
        Me.pnlGenRefundDetails.Controls.Add(Me.lblRefundAmount)
        Me.pnlGenRefundDetails.Controls.Add(Me.lblTotalOverageValue)
        Me.pnlGenRefundDetails.Controls.Add(Me.cmbFromCheck)
        Me.pnlGenRefundDetails.Controls.Add(Me.lblFromCheck)
        Me.pnlGenRefundDetails.Controls.Add(Me.txtComments)
        Me.pnlGenRefundDetails.Controls.Add(Me.lblComments)
        Me.pnlGenRefundDetails.Controls.Add(Me.lblTotalOverage)
        Me.pnlGenRefundDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlGenRefundDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlGenRefundDetails.Name = "pnlGenRefundDetails"
        Me.pnlGenRefundDetails.Size = New System.Drawing.Size(608, 278)
        Me.pnlGenRefundDetails.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblZip)
        Me.GroupBox1.Controls.Add(Me.lblState)
        Me.GroupBox1.Controls.Add(Me.lblCity)
        Me.GroupBox1.Controls.Add(Me.lblAddress2)
        Me.GroupBox1.Controls.Add(Me.lblAddress1)
        Me.GroupBox1.Controls.Add(Me.cboZipCode)
        Me.GroupBox1.Controls.Add(Me.cboState)
        Me.GroupBox1.Controls.Add(Me.cboCity)
        Me.GroupBox1.Controls.Add(Me.txtAddress2)
        Me.GroupBox1.Controls.Add(Me.txtAddress1)
        Me.GroupBox1.Controls.Add(Me.txtName)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.chkOwner)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 80)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(584, 184)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Reimburse Issuer"
        '
        'lblZip
        '
        Me.lblZip.Location = New System.Drawing.Point(408, 149)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(56, 16)
        Me.lblZip.TabIndex = 16
        Me.lblZip.Text = "Zip Code"
        Me.lblZip.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblState
        '
        Me.lblState.Location = New System.Drawing.Point(296, 149)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(32, 16)
        Me.lblState.TabIndex = 17
        Me.lblState.Text = "State"
        Me.lblState.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(16, 144)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(56, 16)
        Me.lblCity.TabIndex = 15
        Me.lblCity.Text = "City"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblAddress2
        '
        Me.lblAddress2.Location = New System.Drawing.Point(16, 112)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(56, 16)
        Me.lblAddress2.TabIndex = 13
        Me.lblAddress2.Text = "Address 2"
        Me.lblAddress2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblAddress1
        '
        Me.lblAddress1.Location = New System.Drawing.Point(16, 80)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New System.Drawing.Size(56, 16)
        Me.lblAddress1.TabIndex = 14
        Me.lblAddress1.Text = "Address 1"
        Me.lblAddress1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cboZipCode
        '
        Me.cboZipCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboZipCode.DisplayMember = ""
        Me.cboZipCode.Location = New System.Drawing.Point(472, 144)
        Me.cboZipCode.Name = "cboZipCode"
        Me.cboZipCode.Size = New System.Drawing.Size(96, 21)
        Me.cboZipCode.TabIndex = 20
        Me.cboZipCode.ValueMember = ""
        '
        'cboState
        '
        Me.cboState.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboState.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Inset
        Me.cboState.DisplayMember = ""
        Me.cboState.Location = New System.Drawing.Point(336, 144)
        Me.cboState.Name = "cboState"
        Me.cboState.Size = New System.Drawing.Size(72, 21)
        Me.cboState.TabIndex = 22
        Me.cboState.ValueMember = ""
        '
        'cboCity
        '
        Me.cboCity.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboCity.DisplayMember = ""
        Me.cboCity.Location = New System.Drawing.Point(88, 144)
        Me.cboCity.Name = "cboCity"
        Me.cboCity.Size = New System.Drawing.Size(200, 21)
        Me.cboCity.TabIndex = 21
        Me.cboCity.ValueMember = ""
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(88, 112)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(200, 20)
        Me.txtAddress2.TabIndex = 19
        Me.txtAddress2.Text = ""
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New System.Drawing.Point(88, 80)
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(200, 20)
        Me.txtAddress1.TabIndex = 18
        Me.txtAddress1.Text = ""
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(88, 48)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(384, 20)
        Me.txtName.TabIndex = 11
        Me.txtName.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Name"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'chkOwner
        '
        Me.chkOwner.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkOwner.Location = New System.Drawing.Point(16, 16)
        Me.chkOwner.Name = "chkOwner"
        Me.chkOwner.Size = New System.Drawing.Size(64, 24)
        Me.chkOwner.TabIndex = 10
        Me.chkOwner.Text = "Owner"
        '
        'txtRefundAmount
        '
        Me.txtRefundAmount.Location = New System.Drawing.Point(492, 16)
        Me.txtRefundAmount.Name = "txtRefundAmount"
        Me.txtRefundAmount.TabIndex = 1
        Me.txtRefundAmount.Text = "0.00"
        '
        'lblRefundAmount
        '
        Me.lblRefundAmount.Location = New System.Drawing.Point(408, 16)
        Me.lblRefundAmount.Name = "lblRefundAmount"
        Me.lblRefundAmount.Size = New System.Drawing.Size(88, 23)
        Me.lblRefundAmount.TabIndex = 7
        Me.lblRefundAmount.Text = "Refund Amount"
        '
        'lblTotalOverageValue
        '
        Me.lblTotalOverageValue.BackColor = System.Drawing.SystemColors.Window
        Me.lblTotalOverageValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalOverageValue.Location = New System.Drawing.Point(96, 16)
        Me.lblTotalOverageValue.Name = "lblTotalOverageValue"
        Me.lblTotalOverageValue.Size = New System.Drawing.Size(64, 17)
        Me.lblTotalOverageValue.TabIndex = 6
        Me.lblTotalOverageValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbFromCheck
        '
        Me.cmbFromCheck.Location = New System.Drawing.Point(256, 16)
        Me.cmbFromCheck.Name = "cmbFromCheck"
        Me.cmbFromCheck.Size = New System.Drawing.Size(136, 21)
        Me.cmbFromCheck.TabIndex = 3
        '
        'lblFromCheck
        '
        Me.lblFromCheck.Location = New System.Drawing.Point(184, 16)
        Me.lblFromCheck.Name = "lblFromCheck"
        Me.lblFromCheck.Size = New System.Drawing.Size(72, 17)
        Me.lblFromCheck.TabIndex = 5
        Me.lblFromCheck.Text = "From Check:"
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(96, 48)
        Me.txtComments.Name = "txtComments"
        Me.txtComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtComments.Size = New System.Drawing.Size(496, 20)
        Me.txtComments.TabIndex = 2
        Me.txtComments.Text = ""
        '
        'lblComments
        '
        Me.lblComments.Location = New System.Drawing.Point(16, 48)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(62, 17)
        Me.lblComments.TabIndex = 3
        Me.lblComments.Text = "Comments:"
        '
        'lblTotalOverage
        '
        Me.lblTotalOverage.Location = New System.Drawing.Point(8, 16)
        Me.lblTotalOverage.Name = "lblTotalOverage"
        Me.lblTotalOverage.Size = New System.Drawing.Size(88, 17)
        Me.lblTotalOverage.TabIndex = 2
        Me.lblTotalOverage.Text = "Check Overage"
        Me.lblTotalOverage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GenerateRefund
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(608, 318)
        Me.Controls.Add(Me.pnlGenRefundDetails)
        Me.Controls.Add(Me.pnlGenerateRefundBottom)
        Me.Name = "GenerateRefund"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Generate Refund"
        Me.pnlGenerateRefundBottom.ResumeLayout(False)
        Me.pnlGenRefundDetails.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.cboZipCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboState, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboCity, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Control Events"
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
#End Region


#Region "Address Manipulation"
    Private Sub cboCity_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCity.Leave
        oInvoice.IssueCity = cboCity.Text
        LoadZipComboBox()
    End Sub

    Private Sub cboState_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboState.Leave
        oInvoice.IssueState = cboState.Text
        LoadCityComboBox()
        LoadZipComboBox()
    End Sub

    Private Sub cboZipCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboZipCode.Leave
        oInvoice.IssueZip = cboZipCode.Text
        LoadCityComboBox()
    End Sub
    Private Sub txtAddress2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddress2.TextChanged
        oInvoice.IssueAddr2 = txtAddress2.Text
    End Sub

    Private Sub txtAddress1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddress1.TextChanged
        oInvoice.IssueAddr1 = txtAddress1.Text
    End Sub


    Private Sub txtName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.TextChanged
        oInvoice.IssueName = txtName.Text
    End Sub

    Private Sub LoadStateComboBox(Optional ByVal strState As String = "")
        Dim strSQL As String

        Try
            strSQL = "SELECT DISTINCT STATE FROM tblSYS_ZIPCODES WHERE"
            strSQL &= " CITY LIKE '%"
            If IsNothing(oInvoice.IssueCity) = False Then
                strSQL &= oInvoice.IssueCity.Trim
            End If
            strSQL &= "%'"
            strSQL &= " AND ZIP LIKE '%"
            If IsNothing(oInvoice.IssueZip) = False Then
                strSQL &= oInvoice.IssueZip.Trim
            End If
            strSQL &= "%'"
            strSQL &= " ORDER BY STATE"

            cboState.DataSource = oAddress.GetDataSet(strSQL)

            If strState <> "" Then
                cboState.Text = strState
                oInvoice.IssueState = strState
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub LoadCityComboBox()
        Dim strSQL As String

        Try
            strSQL = "SELECT DISTINCT CITY FROM tblSYS_ZIPCODES WHERE"
            strSQL &= " STATE LIKE '%"
            If IsNothing(oInvoice.IssueState) = False Then
                strSQL &= oInvoice.IssueState.Trim
            End If
            strSQL &= "%'"
            strSQL &= " AND ZIP LIKE '%"
            If IsNothing(oInvoice.IssueZip) = False Then
                strSQL &= oInvoice.IssueZip.Trim
            End If
            strSQL &= "%'"
            strSQL &= " ORDER BY CITY"

            cboCity.DataSource = oAddress.GetDataSet(strSQL)
        Catch ex As Exception

        End Try

    End Sub
    Public Sub LoadZipComboBox()
        Dim strSQL As String

        Try
            strSQL = "SELECT DISTINCT ZIP FROM tblSYS_ZIPCODES WHERE"
            strSQL &= " STATE LIKE '%"
            If IsNothing(oInvoice.IssueState) = False Then
                strSQL &= oInvoice.IssueState.Trim
            End If
            strSQL &= "%'"
            strSQL &= " AND CITY LIKE '%"
            If IsNothing(oInvoice.IssueCity) = False Then
                strSQL &= oInvoice.IssueCity.Trim
            End If
            strSQL &= "%'"
            strSQL &= " ORDER BY ZIP"

            cboZipCode.DataSource = oAddress.GetDataSet(strSQL)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Sub PopulateAddress()
        Dim oOwner As New MUSTER.BusinessLogic.pOwner
        Dim oAddress As New MUSTER.BusinessLogic.pAddress

        Try
            oOwner.Retrieve(OwnerID)
            oAddress.Retrieve(oOwner.AddressId)

            If oOwner.PersonID > 0 Then
                txtName.Text = oOwner.Persona.FirstName & " " & oOwner.Persona.LastName
            Else
                txtName.Text = oOwner.Organization.Company
            End If


            oInvoice.IssueAddr1 = oAddress.AddressLine1
            oInvoice.IssueAddr2 = oAddress.AddressLine2
            oInvoice.IssueCity = oAddress.City
            oInvoice.IssueState = oAddress.State
            oInvoice.IssueZip = oAddress.Zip

            txtAddress1.Text = oAddress.AddressLine1
            txtAddress2.Text = oAddress.AddressLine2
            cboState.Text = oAddress.State
            If ContainsValue(cboZipCode, oAddress.Zip) Then
                cboZipCode.Text = oAddress.Zip
            Else
                cboZipCode.Text = String.Empty
            End If
            If ContainsValue(cboCity, oAddress.City) Then
                cboCity.Text = oAddress.City
            Else
                cboCity.Text = String.Empty
            End If

            'txtAddress1.Text = oInvoice.IssueAddr1
            'txtAddress2.Text = oInvoice.IssueAddr2
            'cboState.Text = oInvoice.IssueState
            'If ContainsValue(cboCity, oInvoice.IssueCity) Then
            '    cboCity.Text = oInvoice.IssueCity
            'Else
            '    cboCity.Text = String.Empty
            'End If
            'If ContainsValue(cboZipCode, oInvoice.IssueZip) Then
            '    cboZipCode.Text = oInvoice.IssueZip
            'Else
            '    cboZipCode.Text = String.Empty
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function ContainsValue(ByVal cbo As Infragistics.Win.UltraWinGrid.UltraCombo, ByVal value As String) As Boolean
        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In cbo.Rows
            If UCase(row.Cells(0).Value) = UCase(value) Then
                Return True
            End If
        Next
        Return False
    End Function

#End Region



    Private Sub GenerateRefund_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        TotalOverage = oInvoice.GetOverpaymentBucket(OwnerID)
        lblTotalOverageValue.Text = FormatNumber(TotalOverage, 2, TriState.True, TriState.False, TriState.True)

        oInvoice.Retrieve(0)


        bolLoading = True
        LoadStateComboBox("MS")
        LoadCityComboBox()
        LoadZipComboBox()
        LoadCheckComboBox()
        UpdateCheckInfo()

        bolLoading = False

        If Me.cmbFromCheck.DataSource Is Nothing OrElse Me.cmbFromCheck.Items.Count = 0 Then
            MsgBox("There is either no overpaid checks or not enough available overage to apply to a refund.")
            Me.Close()

        End If

    End Sub


    Private Sub chkOwner_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOwner.CheckedChanged
        If chkOwner.Checked Then
            PopulateAddress()
        End If
    End Sub


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim FiscalYear As Int16

        Try
            If txtRefundAmount.Text = "" Or Not (txtRefundAmount.Text > "0") Then
                MsgBox("Refund Amount Must Be Greater Than Zero", MsgBoxStyle.OKOnly, "Invalid Refund Amount")
                Exit Sub
            End If
            If TotalOverage < txtRefundAmount.Text Then
                MsgBox("Refund Amount Can Not Exceed Total Overage", MsgBoxStyle.OKOnly, "Invalid Refund Amount")
                Exit Sub
            End If
            If Len(Trim(txtComments.Text)) = 0 Then
                MsgBox("Refund Comments Can Not Be Blank", MsgBoxStyle.OKOnly, "Invalid Refund Comment")
                Exit Sub
            End If

            If Len(Trim(txtName.Text)) = 0 Then
                MsgBox("Issuer Name Can Not Be Blank", MsgBoxStyle.OKOnly, "Invalid String")
                Exit Sub
            End If
            If Len(Trim(txtAddress1.Text)) = 0 Then
                MsgBox("Issuer Address Line 1 Can Not Be Blank", MsgBoxStyle.OKOnly, "Invalid String")
                Exit Sub
            End If
            If Len(Trim(cboCity.Text)) = 0 Then
                MsgBox("Issuer City Can Not Be Blank", MsgBoxStyle.OKOnly, "Invalid String")
                Exit Sub
            End If
            If Len(Trim(cboState.Text)) = 0 Then
                MsgBox("Issuer State Can Not Be Blank", MsgBoxStyle.OKOnly, "Invalid String")
                Exit Sub
            End If
            If Len(Trim(cboZipCode.Text)) = 0 Then
                MsgBox("Issuer Zip Code Can Not Be Blank", MsgBoxStyle.OKOnly, "Invalid String")
                Exit Sub
            End If

            oInvoice.InvoiceType = "R"
            oInvoice.RecType = "ADVIC"
            oInvoice.FeeType = "NA"
            oInvoice.FiscalYear = FiscalYear
            oInvoice.OwnerID = OwnerID
            oInvoice.InvoiceAmount = txtRefundAmount.Text

            ' Get Fiscal Year from Fees Basis
            FiscalYear = oFeeBasis.GetFiscalYear(Now.Date)
            'FiscalYear = DatePart(DateInterval.Year, Now.Date)
            'If DatePart(DateInterval.Month, Now.Date) > 6 Then
            '    FiscalYear += 1
            'End If

            oInvoiceADLN.InvoiceType = "R"
            oInvoiceADLN.FeeType = "R"
            oInvoiceADLN.RecType = "ADLN"
            oInvoiceADLN.FiscalYear = FiscalYear
            oInvoiceADLN.OwnerID = OwnerID
            oInvoiceADLN.FacilityID = 0
            oInvoiceADLN.SequenceNumber = 1
            oInvoiceADLN.Description = txtComments.Text
            oInvoiceADLN.InvoiceLineAmount = txtRefundAmount.Text
            oInvoiceADLN.Quantity = 1
            oInvoiceADLN.UnitPrice = txtRefundAmount.Text
            oInvoice.InvoiceLineItems.Add(oInvoiceADLN)

            If oInvoice.ID <= 0 Then
                oInvoice.CreatedBy = MusterContainer.AppUser.ID
            Else
                oInvoice.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oInvoice.SaveNewInvoice(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If


            MsgBox("Refund Request Saved")
            Me.Close()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot btnSave_Click " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub txtRefundAmount_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRefundAmount.Leave
        If IsNumeric(txtRefundAmount.Text) Then

            txtRefundAmount.Text = FormatNumber(txtRefundAmount.Text, 2, TriState.True, TriState.False, TriState.True)
            oInvoice.InvoiceAmount = txtRefundAmount.Text

        Else
            MsgBox("Refund Amount Must Be Numeric.")

        End If
    End Sub

    Private Sub cmbFromCheck_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFromCheck.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oInvoice.CheckNumber = cmbFromCheck.Text
        oInvoice.CheckTransID = cmbFromCheck.SelectedValue
        UpdateCheckInfo()

    End Sub

    Private Sub LoadCheckComboBox()
        Dim strSQL As String
        Dim dsLocal As New DataSet

        Try
            strSQL = "SELECT CHECK_TRANS_ID as CheckID, CHECK_NUMBER as CheckNum FROM vFees_RefundableChecks "
            strSQL &= " Where Owner_ID = " & OwnerID
            strSQL &= " and isnull((select OVERPAYMENTBUCKET  from vFees_OverPayment_bucket where OWNER_ID = " & OwnerID & " ),0) <= "
            strSQL &= " (select sum(AvailableOverpayment) FROM vFees_RefundableChecks Where Owner_ID =" & OwnerID & ") "
            strSQL &= " ORDER BY CHECK_NUMBER"
            dsLocal = oAddress.GetDataSet(strSQL)

            cmbFromCheck.DataSource = dsLocal.Tables(0)
            cmbFromCheck.DisplayMember = "CheckNum"
            cmbFromCheck.ValueMember = "CheckID"


            If cmbFromCheck.Items.Count > 0 Then
                cmbFromCheck.SelectedIndex = 0
                oInvoice.CheckTransID = cmbFromCheck.SelectedValue
                oInvoice.CheckNumber = cmbFromCheck.Text
            End If


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot LoadCheckComboBox " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub UpdateCheckInfo()
        Dim CheckOverageAmt As Single
        'TotalOverage As Single
        Try
            If cmbFromCheck.Items.Count = 0 Then Exit Sub

            If chkOwner.Checked = False Then
                Me.txtName.Text = oInvoice.GetCheckIssuerName(OwnerID, cmbFromCheck.SelectedValue, cmbFromCheck.Text)
            End If
            CheckOverageAmt = oInvoice.GetRefundableCheckAmount(OwnerID, cmbFromCheck.SelectedValue, cmbFromCheck.Text)

            If CheckOverageAmt < TotalOverage Then
                lblTotalOverageValue.Text = FormatNumber(CheckOverageAmt, 2, TriState.True, TriState.False, TriState.True)
                txtRefundAmount.Text = lblTotalOverageValue.Text
            Else
                lblTotalOverageValue.Text = FormatNumber(TotalOverage, 2, TriState.True, TriState.False, TriState.True)
                txtRefundAmount.Text = lblTotalOverageValue.Text
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot UpdateCheckInfo " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
