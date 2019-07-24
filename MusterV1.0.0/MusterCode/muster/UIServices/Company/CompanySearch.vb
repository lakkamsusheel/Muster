Public Class CompanySearch
    Inherits System.Windows.Forms.Form
#Region " USER DEFINED VARIABLES"
    Dim pCompany As New MUSTER.BusinessLogic.pCompany
    Dim strModule As String = String.Empty
#End Region
#Region " Events"
    Public Event LicenseeCompanyDetails(ByVal licensee_id As Integer, ByVal company_id As Integer, ByVal LicenseeName As String, ByVal companyName As String)
    Public Event CompanyDetails(ByVal company_id As Integer, ByVal companyName As String)
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    'P1 04/03/06 new
    Public Sub New(ByVal strMod As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        strModule = strMod
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
    Friend WithEvents lblSearchResults As System.Windows.Forms.Label
    Friend WithEvents ugSearchResults As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents chkIRAC As System.Windows.Forms.CheckBox
    Friend WithEvents chkERAC As System.Windows.Forms.CheckBox
    Friend WithEvents txtCompanyName As System.Windows.Forms.TextBox
    Friend WithEvents lblCompanyName As System.Windows.Forms.Label
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress As System.Windows.Forms.TextBox
    Friend WithEvents txtLicenseeName As System.Windows.Forms.TextBox
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents lblLicenseeName As System.Windows.Forms.Label
    Friend WithEvents cmbState As System.Windows.Forms.ComboBox
    Friend WithEvents lblState As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblSearchResults = New System.Windows.Forms.Label
        Me.ugSearchResults = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkIRAC = New System.Windows.Forms.CheckBox
        Me.chkERAC = New System.Windows.Forms.CheckBox
        Me.txtCompanyName = New System.Windows.Forms.TextBox
        Me.lblCompanyName = New System.Windows.Forms.Label
        Me.btnClear = New System.Windows.Forms.Button
        Me.btnSearch = New System.Windows.Forms.Button
        Me.txtCity = New System.Windows.Forms.TextBox
        Me.txtAddress = New System.Windows.Forms.TextBox
        Me.txtLicenseeName = New System.Windows.Forms.TextBox
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblAddress = New System.Windows.Forms.Label
        Me.lblLicenseeName = New System.Windows.Forms.Label
        Me.cmbState = New System.Windows.Forms.ComboBox
        Me.lblState = New System.Windows.Forms.Label
        CType(Me.ugSearchResults, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblSearchResults
        '
        Me.lblSearchResults.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSearchResults.Location = New System.Drawing.Point(8, 240)
        Me.lblSearchResults.Name = "lblSearchResults"
        Me.lblSearchResults.Size = New System.Drawing.Size(96, 16)
        Me.lblSearchResults.TabIndex = 164
        Me.lblSearchResults.Text = "SearchResults"
        '
        'ugSearchResults
        '
        Me.ugSearchResults.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugSearchResults.Location = New System.Drawing.Point(8, 256)
        Me.ugSearchResults.Name = "ugSearchResults"
        Me.ugSearchResults.Size = New System.Drawing.Size(688, 208)
        Me.ugSearchResults.TabIndex = 163
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(232, 472)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 23)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(152, 472)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(64, 23)
        Me.btnOk.TabIndex = 9
        Me.btnOk.Text = "OK"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkIRAC)
        Me.GroupBox1.Controls.Add(Me.chkERAC)
        Me.GroupBox1.Controls.Add(Me.txtCompanyName)
        Me.GroupBox1.Controls.Add(Me.lblCompanyName)
        Me.GroupBox1.Controls.Add(Me.btnClear)
        Me.GroupBox1.Controls.Add(Me.btnSearch)
        Me.GroupBox1.Controls.Add(Me.txtCity)
        Me.GroupBox1.Controls.Add(Me.txtAddress)
        Me.GroupBox1.Controls.Add(Me.txtLicenseeName)
        Me.GroupBox1.Controls.Add(Me.lblCity)
        Me.GroupBox1.Controls.Add(Me.lblAddress)
        Me.GroupBox1.Controls.Add(Me.lblLicenseeName)
        Me.GroupBox1.Controls.Add(Me.cmbState)
        Me.GroupBox1.Controls.Add(Me.lblState)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(336, 232)
        Me.GroupBox1.TabIndex = 173
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Search Parameters:"
        '
        'chkIRAC
        '
        Me.chkIRAC.Location = New System.Drawing.Point(208, 160)
        Me.chkIRAC.Name = "chkIRAC"
        Me.chkIRAC.Size = New System.Drawing.Size(64, 24)
        Me.chkIRAC.TabIndex = 6
        Me.chkIRAC.Text = "IRAC"
        '
        'chkERAC
        '
        Me.chkERAC.Location = New System.Drawing.Point(128, 160)
        Me.chkERAC.Name = "chkERAC"
        Me.chkERAC.Size = New System.Drawing.Size(56, 24)
        Me.chkERAC.TabIndex = 5
        Me.chkERAC.Text = "ERAC"
        '
        'txtCompanyName
        '
        Me.txtCompanyName.Location = New System.Drawing.Point(128, 56)
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.Size = New System.Drawing.Size(144, 20)
        Me.txtCompanyName.TabIndex = 1
        Me.txtCompanyName.Text = ""
        '
        'lblCompanyName
        '
        Me.lblCompanyName.Location = New System.Drawing.Point(24, 56)
        Me.lblCompanyName.Name = "lblCompanyName"
        Me.lblCompanyName.Size = New System.Drawing.Size(96, 16)
        Me.lblCompanyName.TabIndex = 184
        Me.lblCompanyName.Text = "Company Name:"
        Me.lblCompanyName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(208, 192)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(64, 23)
        Me.btnClear.TabIndex = 8
        Me.btnClear.Text = "Clear"
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(128, 192)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(64, 23)
        Me.btnSearch.TabIndex = 7
        Me.btnSearch.Text = "Search"
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(128, 104)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(144, 20)
        Me.txtCity.TabIndex = 3
        Me.txtCity.Text = ""
        '
        'txtAddress
        '
        Me.txtAddress.Location = New System.Drawing.Point(128, 80)
        Me.txtAddress.Name = "txtAddress"
        Me.txtAddress.Size = New System.Drawing.Size(144, 20)
        Me.txtAddress.TabIndex = 2
        Me.txtAddress.Text = ""
        '
        'txtLicenseeName
        '
        Me.txtLicenseeName.Location = New System.Drawing.Point(128, 24)
        Me.txtLicenseeName.Name = "txtLicenseeName"
        Me.txtLicenseeName.Size = New System.Drawing.Size(144, 20)
        Me.txtLicenseeName.TabIndex = 0
        Me.txtLicenseeName.Text = ""
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(80, 104)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(40, 16)
        Me.lblCity.TabIndex = 180
        Me.lblCity.Text = "City:"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(8, 80)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(112, 16)
        Me.lblAddress.TabIndex = 179
        Me.lblAddress.Text = "Licensee Address:"
        Me.lblAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLicenseeName
        '
        Me.lblLicenseeName.Location = New System.Drawing.Point(32, 24)
        Me.lblLicenseeName.Name = "lblLicenseeName"
        Me.lblLicenseeName.Size = New System.Drawing.Size(88, 16)
        Me.lblLicenseeName.TabIndex = 178
        Me.lblLicenseeName.Text = "Licensee Name:"
        Me.lblLicenseeName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbState
        '
        Me.cmbState.Location = New System.Drawing.Point(128, 128)
        Me.cmbState.Name = "cmbState"
        Me.cmbState.Size = New System.Drawing.Size(56, 21)
        Me.cmbState.TabIndex = 4
        '
        'lblState
        '
        Me.lblState.Location = New System.Drawing.Point(72, 128)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(48, 16)
        Me.lblState.TabIndex = 177
        Me.lblState.Text = "State:"
        Me.lblState.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CompanySearch
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(704, 502)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.lblSearchResults)
        Me.Controls.Add(Me.ugSearchResults)
        Me.Name = "CompanySearch"
        Me.Text = "CompanySearch"
        CType(Me.ugSearchResults, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim dsSearchResult As DataSet
        Try
            If strModule <> String.Empty And strModule = "Technical" Then
                dsSearchResult = pCompany.searchTecCompany(txtCompanyName.Text, txtAddress.Text, txtCity.Text, cmbState.Text, chkERAC.Checked, chkIRAC.Checked)
                dsSearchResult.Tables(0).DefaultView.Sort = "COMPANY_NAME ASC"
                ' dsSearchResult.Tables(0).Columns("DELETED").ColumnName = "PRIOR ASSOCIATION"
                ugSearchResults.DataSource = dsSearchResult.Tables(0).DefaultView
            Else
                dsSearchResult = pCompany.searchLicensee(txtLicenseeName.Text, txtCompanyName.Text, txtAddress.Text, txtCity.Text, cmbState.Text, chkERAC.Checked, chkIRAC.Checked, "spCOMSearch")
                dsSearchResult.Tables(0).DefaultView.Sort = "LICENSEE_NAME ASC"
                dsSearchResult.Tables(0).Columns("DELETED").ColumnName = "PRIOR ASSOCIATION"
                ugSearchResults.DataSource = dsSearchResult.Tables(0).DefaultView
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Try
            If Not ugSearchResults.ActiveRow Is Nothing Then
                If strModule <> String.Empty And strModule = "Technical" Then
                    RaiseEvent CompanyDetails(ugSearchResults.ActiveRow.Cells("COMPANY_ID").Value, _
                                              ugSearchResults.ActiveRow.Cells("COMPANY_NAME").Value)
                Else
                    RaiseEvent LicenseeCompanyDetails(IIf(ugSearchResults.ActiveRow.Cells("LICENSEE_ID").Value Is System.DBNull.Value, 0, ugSearchResults.ActiveRow.Cells("LICENSEE_ID").Value), _
                                   ugSearchResults.ActiveRow.Cells("COMPANY_ID").Value, _
                                   IIf(ugSearchResults.ActiveRow.Cells("LICENSEE_NAME").Value Is System.DBNull.Value, "", ugSearchResults.ActiveRow.Cells("LICENSEE_NAME").Value), _
                                   ugSearchResults.ActiveRow.Cells("COMPANY_NAME").Value)
                End If
            Else
                MsgBox("Please select a row from the grid")
            End If
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub ugSearchResults_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugSearchResults.InitializeLayout
        If strModule <> String.Empty And strModule = "Technical" Then
            ugSearchResults.DisplayLayout.Bands(0).Columns("COMPANY_ID").Hidden = True
            'ugSearchResults.DisplayLayout.Bands(0).Columns("LICENSEE_ID").Hidden = True
            ugSearchResults.DisplayLayout.Bands(0).Columns("ADDRESS").Width = 300
            ugSearchResults.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        Else
            ugSearchResults.DisplayLayout.Bands(0).Columns("COMPANY_ID").Hidden = True
            ugSearchResults.DisplayLayout.Bands(0).Columns("LICENSEE_ID").Hidden = True
            ugSearchResults.DisplayLayout.Bands(0).Columns("ADDRESS").Width = 300
            ugSearchResults.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        End If

    End Sub

    'Private Sub ugSearchResults_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugSearchResults.DoubleClick
    '    Try
    '        If Not ugSearchResults.ActiveRow Is Nothing Then
    '            RaiseEvent LicenseeCompanyDetails(ugSearchResults.ActiveRow.Cells("LICENSEE_ID").Value, _
    '            ugSearchResults.ActiveRow.Cells("COMPANY_ID").Value, _
    '            ugSearchResults.ActiveRow.Cells("LICENSEE_NAME").Value, _
    '            ugSearchResults.ActiveRow.Cells("COMPANY_NAME").Value)
    '            MsgBox("Licensee is selected successfully")
    '            Me.Close()
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            Me.txtLicenseeName.Text = String.Empty
            Me.txtCompanyName.Text = String.Empty
            Me.txtAddress.Text = String.Empty
            Me.txtCity.Text = String.Empty
            Me.cmbState.Text = String.Empty
            Me.chkERAC.Checked = False
            Me.chkIRAC.Checked = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub CompanySearch_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If strModule <> String.Empty And strModule = "Technical" Then
                txtLicenseeName.Visible = False
                lblLicenseeName.Visible = False
            Else
                txtLicenseeName.Visible = True
                lblLicenseeName.Visible = True
                txtLicenseeName.Focus()
            End If
            'Me.Activate()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub CompanySearch_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        If txtLicenseeName.CanFocus() Then
            txtLicenseeName.Focus()
        End If
    End Sub

    ' #1463 Begin
    Private Sub txtLicenseeName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtLicenseeName.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnSearch.PerformClick()
        End If
    End Sub
    Private Sub txtCompanyName_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCompanyName.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnSearch.PerformClick()
        End If
    End Sub
    Private Sub txtAddress_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddress.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnSearch.PerformClick()
        End If
    End Sub
    Private Sub txtCity_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCity.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnSearch.PerformClick()
        End If
    End Sub
    Private Sub cmbState_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbState.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            btnSearch.PerformClick()
        End If
    End Sub
    ' #1463 End

End Class
