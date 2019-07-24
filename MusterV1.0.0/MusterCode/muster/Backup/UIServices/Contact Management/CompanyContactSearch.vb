Public Class CompanyContactSearch

    Inherits System.Windows.Forms.Form
    Private pConStruct As MUSTER.BusinessLogic.pContactStruct
    Private WithEvents ContactFrm As Contacts
    Dim nEntityID As Integer = 0
    Dim nEntityTypeID As Integer = 0
    Dim strModuleName As String = String.Empty
    Dim nContactID As Integer = 0
    Dim result As DialogResult

    Friend Event NewCompanyAdded()
    Private frmContact As Contacts
    Dim returnVal As String = String.Empty

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByVal ContactID As Integer, ByVal EntityID As Integer, ByVal EntityTypeID As Integer, ByVal ModuleName As String, Optional ByRef pContactStruct As MUSTER.BusinessLogic.pContactStruct = Nothing, Optional ByRef frm As Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        nContactID = ContactID
        pConStruct = pContactStruct
        nEntityID = EntityID
        nEntityTypeID = EntityTypeID
        strModuleName = ModuleName
        frmContact = CType(frm, Contacts)

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
    Friend WithEvents cmbState As System.Windows.Forms.ComboBox
    Friend WithEvents lblState As System.Windows.Forms.Label
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents lblContactName As System.Windows.Forms.Label
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtContactName As System.Windows.Forms.TextBox
    Friend WithEvents lblOrganization As System.Windows.Forms.Label
    Friend WithEvents lblCell As System.Windows.Forms.Label
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Friend WithEvents lblPhone1 As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblSearchResults As System.Windows.Forms.Label
    Friend WithEvents ugSearchResults As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnAddNewContact As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnAssociate As System.Windows.Forms.Button
    Friend WithEvents grpSearch As System.Windows.Forms.GroupBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Public WithEvents mskTxtFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtCell As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtPhone2 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtPhone1 As AxMSMask.AxMaskEdBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CompanyContactSearch))
        Me.cmbState = New System.Windows.Forms.ComboBox
        Me.lblCell = New System.Windows.Forms.Label
        Me.lblPhone2 = New System.Windows.Forms.Label
        Me.lblState = New System.Windows.Forms.Label
        Me.lblPhone1 = New System.Windows.Forms.Label
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblAddress = New System.Windows.Forms.Label
        Me.lblContactName = New System.Windows.Forms.Label
        Me.txtCity = New System.Windows.Forms.TextBox
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.txtContactName = New System.Windows.Forms.TextBox
        Me.lblOrganization = New System.Windows.Forms.Label
        Me.lblEmail = New System.Windows.Forms.Label
        Me.lblFax = New System.Windows.Forms.Label
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.lblSearchResults = New System.Windows.Forms.Label
        Me.ugSearchResults = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnAddNewContact = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnAssociate = New System.Windows.Forms.Button
        Me.grpSearch = New System.Windows.Forms.GroupBox
        Me.btnSearch = New System.Windows.Forms.Button
        Me.btnClear = New System.Windows.Forms.Button
        Me.mskTxtFax = New AxMSMask.AxMaskEdBox
        Me.mskTxtCell = New AxMSMask.AxMaskEdBox
        Me.mskTxtPhone2 = New AxMSMask.AxMaskEdBox
        Me.mskTxtPhone1 = New AxMSMask.AxMaskEdBox
        CType(Me.ugSearchResults, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpSearch.SuspendLayout()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbState
        '
        Me.cmbState.Location = New System.Drawing.Point(120, 104)
        Me.cmbState.Name = "cmbState"
        Me.cmbState.Size = New System.Drawing.Size(56, 21)
        Me.cmbState.TabIndex = 3
        '
        'lblCell
        '
        Me.lblCell.Location = New System.Drawing.Point(536, 32)
        Me.lblCell.Name = "lblCell"
        Me.lblCell.Size = New System.Drawing.Size(32, 16)
        Me.lblCell.TabIndex = 24
        Me.lblCell.Text = "Cell:"
        Me.lblCell.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPhone2
        '
        Me.lblPhone2.Location = New System.Drawing.Point(288, 56)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone2.TabIndex = 23
        Me.lblPhone2.Text = "Phone 2:"
        Me.lblPhone2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblState
        '
        Me.lblState.Location = New System.Drawing.Point(64, 104)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(48, 16)
        Me.lblState.TabIndex = 22
        Me.lblState.Text = "State:"
        Me.lblState.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPhone1
        '
        Me.lblPhone1.Location = New System.Drawing.Point(288, 32)
        Me.lblPhone1.Name = "lblPhone1"
        Me.lblPhone1.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone1.TabIndex = 21
        Me.lblPhone1.Text = "Phone 1:"
        Me.lblPhone1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(72, 80)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(40, 16)
        Me.lblCity.TabIndex = 30
        Me.lblCity.Text = "City:"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(56, 56)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(56, 16)
        Me.lblAddress.TabIndex = 29
        Me.lblAddress.Text = "Address:"
        Me.lblAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblContactName
        '
        Me.lblContactName.Location = New System.Drawing.Point(24, 32)
        Me.lblContactName.Name = "lblContactName"
        Me.lblContactName.Size = New System.Drawing.Size(88, 16)
        Me.lblContactName.TabIndex = 28
        Me.lblContactName.Text = "Contact Name:"
        Me.lblContactName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(120, 80)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(144, 20)
        Me.txtCity.TabIndex = 2
        Me.txtCity.Text = ""
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(120, 56)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(144, 20)
        Me.txtAddress.TabIndex = 1
        Me.txtAddress.Text = ""
        '
        'txtContactName
        '
        Me.txtContactName.Location = New System.Drawing.Point(120, 32)
        Me.txtContactName.Name = "txtContactName"
        Me.txtContactName.Size = New System.Drawing.Size(144, 20)
        Me.txtContactName.TabIndex = 0
        Me.txtContactName.Text = ""
        '
        'lblOrganization
        '
        Me.lblOrganization.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrganization.ForeColor = System.Drawing.Color.DarkBlue
        Me.lblOrganization.Location = New System.Drawing.Point(24, 8)
        Me.lblOrganization.Name = "lblOrganization"
        Me.lblOrganization.Size = New System.Drawing.Size(88, 16)
        Me.lblOrganization.TabIndex = 31
        Me.lblOrganization.Text = "Organization:"
        Me.lblOrganization.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(520, 80)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(48, 16)
        Me.lblEmail.TabIndex = 35
        Me.lblEmail.Text = "E-mail:"
        Me.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFax
        '
        Me.lblFax.Location = New System.Drawing.Point(528, 56)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(40, 16)
        Me.lblFax.TabIndex = 34
        Me.lblFax.Text = "Fax:"
        Me.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(576, 80)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(144, 20)
        Me.txtEmail.TabIndex = 8
        Me.txtEmail.Text = ""
        '
        'lblSearchResults
        '
        Me.lblSearchResults.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchResults.Location = New System.Drawing.Point(8, 168)
        Me.lblSearchResults.Name = "lblSearchResults"
        Me.lblSearchResults.Size = New System.Drawing.Size(96, 16)
        Me.lblSearchResults.TabIndex = 132
        Me.lblSearchResults.Text = "SearchResults"
        '
        'ugSearchResults
        '
        Me.ugSearchResults.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugSearchResults.Location = New System.Drawing.Point(8, 192)
        Me.ugSearchResults.Name = "ugSearchResults"
        Me.ugSearchResults.Size = New System.Drawing.Size(832, 272)
        Me.ugSearchResults.TabIndex = 11
        '
        'btnAddNewContact
        '
        Me.btnAddNewContact.Location = New System.Drawing.Point(408, 480)
        Me.btnAddNewContact.Name = "btnAddNewContact"
        Me.btnAddNewContact.Size = New System.Drawing.Size(112, 26)
        Me.btnAddNewContact.TabIndex = 14
        Me.btnAddNewContact.Text = "Add New Company"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(328, 480)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(70, 26)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Cancel"
        '
        'btnAssociate
        '
        Me.btnAssociate.Location = New System.Drawing.Point(248, 480)
        Me.btnAssociate.Name = "btnAssociate"
        Me.btnAssociate.Size = New System.Drawing.Size(70, 26)
        Me.btnAssociate.TabIndex = 12
        Me.btnAssociate.Text = "Associate"
        '
        'grpSearch
        '
        Me.grpSearch.Controls.Add(Me.btnSearch)
        Me.grpSearch.Controls.Add(Me.btnClear)
        Me.grpSearch.Location = New System.Drawing.Point(304, 112)
        Me.grpSearch.Name = "grpSearch"
        Me.grpSearch.Size = New System.Drawing.Size(192, 52)
        Me.grpSearch.TabIndex = 180
        Me.grpSearch.TabStop = False
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(24, 16)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(72, 26)
        Me.btnSearch.TabIndex = 12
        Me.btnSearch.Text = "Search"
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(104, 16)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(72, 26)
        Me.btnClear.TabIndex = 11
        Me.btnClear.Text = "Clear"
        '
        'mskTxtFax
        '
        Me.mskTxtFax.Location = New System.Drawing.Point(576, 56)
        Me.mskTxtFax.Name = "mskTxtFax"
        Me.mskTxtFax.OcxState = CType(resources.GetObject("mskTxtFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFax.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtFax.TabIndex = 189
        '
        'mskTxtCell
        '
        Me.mskTxtCell.Location = New System.Drawing.Point(576, 32)
        Me.mskTxtCell.Name = "mskTxtCell"
        Me.mskTxtCell.OcxState = CType(resources.GetObject("mskTxtCell.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtCell.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtCell.TabIndex = 188
        '
        'mskTxtPhone2
        '
        Me.mskTxtPhone2.Location = New System.Drawing.Point(352, 56)
        Me.mskTxtPhone2.Name = "mskTxtPhone2"
        Me.mskTxtPhone2.OcxState = CType(resources.GetObject("mskTxtPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone2.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtPhone2.TabIndex = 187
        '
        'mskTxtPhone1
        '
        Me.mskTxtPhone1.Location = New System.Drawing.Point(352, 32)
        Me.mskTxtPhone1.Name = "mskTxtPhone1"
        Me.mskTxtPhone1.OcxState = CType(resources.GetObject("mskTxtPhone1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone1.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtPhone1.TabIndex = 186
        '
        'CompanyContactSearch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(856, 518)
        Me.Controls.Add(Me.mskTxtFax)
        Me.Controls.Add(Me.mskTxtCell)
        Me.Controls.Add(Me.mskTxtPhone2)
        Me.Controls.Add(Me.mskTxtPhone1)
        Me.Controls.Add(Me.grpSearch)
        Me.Controls.Add(Me.btnAddNewContact)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnAssociate)
        Me.Controls.Add(Me.lblSearchResults)
        Me.Controls.Add(Me.ugSearchResults)
        Me.Controls.Add(Me.lblEmail)
        Me.Controls.Add(Me.lblFax)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.txtCity)
        Me.Controls.Add(Me.txtAddress)
        Me.Controls.Add(Me.txtContactName)
        Me.Controls.Add(Me.lblOrganization)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.lblContactName)
        Me.Controls.Add(Me.cmbState)
        Me.Controls.Add(Me.lblCell)
        Me.Controls.Add(Me.lblPhone2)
        Me.Controls.Add(Me.lblState)
        Me.Controls.Add(Me.lblPhone1)
        Me.Name = "CompanyContactSearch"
        Me.Text = "Company Contact Search"
        CType(Me.ugSearchResults, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpSearch.ResumeLayout(False)
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ContactFrm_NewCompanyAdded() Handles ContactFrm.NewCompanyAdded
        RaiseEvent NewCompanyAdded()
        Me.Close()
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim oContactStruct As New MUSTER.BusinessLogic.pContactStruct
        Try
            Dim dsSearchResult As DataSet
            dsSearchResult = oContactStruct.SearchContact(txtContactName.Text, txtAddress.Text, txtCity.Text, cmbState.Text, _
                            IIf(mskTxtPhone1.FormattedText.ToString = "(___)___-____", "", mskTxtPhone1.FormattedText), _
                            IIf(mskTxtPhone2.FormattedText.ToString = "(___)___-____", "", mskTxtPhone2.FormattedText), _
                            IIf(mskTxtCell.FormattedText.ToString = "(___)___-____", "", mskTxtCell.FormattedText), _
                            IIf(mskTxtFax.FormattedText.ToString = "(___)___-____", "", mskTxtFax.FormattedText), txtEmail.Text, "spCONSearchCompany")
            ugSearchResults.DataSource = dsSearchResult.Tables(0).DefaultView

            '------ rename column headings -----------------------------------------
            ugSearchResults.DisplayLayout.Bands(0).Columns("Company_Name").Header.Caption = "Company Name"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Address_One").Header.Caption = "Address 1"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Address_Two").Header.Caption = "Address 2"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Company_name").Header.Caption = "Company Name"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Phone_Number_One").Header.Caption = "Phone 1"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Ext_One").Header.Caption = "Ext 1"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Phone_Number_Two").Header.Caption = "Phone 2"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Ext_Two").Header.Caption = "Ext 2"
            ugSearchResults.DisplayLayout.Bands(0).Columns("DATE_CREATED").Header.Caption = "Date Created"
            ugSearchResults.DisplayLayout.Bands(0).Columns("CREATED_BY").Header.Caption = "Created By"
            ugSearchResults.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Header.Caption = "Last Edited"
            ugSearchResults.DisplayLayout.Bands(0).Columns("LAST_EDITED_BY").Header.Caption = "Edited By"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Cell_Number").Header.Caption = "Cell Number"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Fax_Number").Header.Caption = "Fax Number"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Email_Address").Header.Caption = "Email Address"
            ugSearchResults.DisplayLayout.Bands(0).Columns("Email_Address_Personal").Header.Caption = "Personal Email"


            ugSearchResults.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
            ugSearchResults.DisplayLayout.Bands(0).Columns("IsPerson").Hidden = True
            ugSearchResults.DisplayLayout.Bands(0).Columns("Org_Entity_Code").Hidden = True
            ugSearchResults.DisplayLayout.Bands(0).Columns("FIPS_Code").Hidden = True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            ClearSearchForm()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAssociate.Click
        Try
            If ugSearchResults.Rows.Count <= 0 Then Exit Sub

            If ugSearchResults.ActiveRow Is Nothing Then
                MsgBox("Select row to Associate.")
                Exit Sub
            End If
            result = MessageBox.Show("Are you sure you want to Associate?", "Contact", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.No Then
                Exit Sub
            End If

            '----- if the Copy Company Address to Person Contact Address checkbox is checked then 
            '      create a new Person Contact record ...with the Company Address ... and 
            '      associate this Company Address to the new Person Contact record 
            Dim msgResult As MsgBoxResult
            msgResult = MsgBox("Do you want the Person Contact's address to be the same as the Company address?", MsgBoxStyle.YesNo, "Contact Management")
            If msgResult = MsgBoxResult.Yes Then
                'pConStruct.bolCopyCompanyAddress = True
                frmContact.txtAddressOne.Text = ugSearchResults.ActiveRow.Cells("Address_One").Value
                frmContact.txtAddressTwo.Text = ugSearchResults.ActiveRow.Cells("Address_Two").Value
                frmContact.cmbState.Text = ugSearchResults.ActiveRow.Cells("State").Value
                frmContact.cmbCity.Text = ugSearchResults.ActiveRow.Cells("City").Value
                'frmContact.mskTxtZip.SelText = ugSearchResults.ActiveRow.Cells("ZipCode").Value
                frmContact.CmbZip.Text = ugSearchResults.ActiveRow.Cells("ZipCode").Value
                frmContact.bolAddChangedFromCompany = True
                pConStruct.childContactID = ugSearchResults.ActiveRow.Cells("ContactID").Value
                pConStruct.contactStructInfo.ParentContactID = nContactID
                'pConStruct.Save(CType(UIUtilsGen.ModuleID.ContactManagement, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, "SEARCH")
                pConStruct.Save(CType(UIUtilsGen.ModuleID.ContactManagement, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, "SEARCH", , , , , , , , True)
                frmContact.strMode = "ASSOCIATE NEW"
                MsgBox("Successfully Associated")
            Else
                'pConStruct.bolCopyCompanyAddress = False
                pConStruct.childContactID = ugSearchResults.ActiveRow.Cells("ContactID").Value
                pConStruct.contactStructInfo.ParentContactID = nContactID
                'pConStruct.Save(CType(UIUtilsGen.ModuleID.ContactManagement, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, "SEARCH")
                pConStruct.Save(CType(UIUtilsGen.ModuleID.ContactManagement, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, "SEARCH", , , , , , , , True)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                frmContact.strMode = "ASSOCIATE NEW"
                MsgBox("Successfully Associated")
            End If

            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            If ex.Message = "THE ASSOCIATION WAS ALREADY DONE BEFORE" Then
                MsgBox(ex.Message)
                Me.Close()
            Else
                MyErr.ShowDialog()
            End If
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnAddNewContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNewContact.Click
        Try
            'If IsNothing(ContactFrm) Then
            '    ContactFrm = New Contacts(nEntityID, nEntityTypeID, strModuleName, 0, Nothing, pConStruct, "ASSOCIATE NEW COMPANY", , True)
            'End If
            'ContactFrm.Show()
            If Not IsNothing(ContactFrm) Then
                ContactFrm = Nothing
            End If
            ContactFrm = New Contacts(nEntityID, nEntityTypeID, strModuleName, 0, Nothing, pConStruct, "ASSOCIATE NEW COMPANY", , True)
            ContactFrm.Show()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ClearSearchForm()
        txtContactName.Text = String.Empty
        txtAddress.Text = String.Empty
        txtCity.Text = String.Empty
        cmbState.Text = String.Empty
        mskTxtPhone1.SelText = String.Empty
        mskTxtPhone2.SelText = String.Empty
        mskTxtCell.SelText = String.Empty
        mskTxtFax.SelText = String.Empty
        txtEmail.Text = String.Empty
    End Sub

    Private Sub CompanyContactSearch_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            cmbState.DisplayMember = "STATE"
            cmbState.DataSource = pConStruct.getStates.Tables(0)
            cmbState.SelectedIndex = -1
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub chkBoxCopyCompanyAddress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

End Class
