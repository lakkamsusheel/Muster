Public Class CompanySearchResults
    Inherits System.Windows.Forms.Form
    Friend WithEvents frmRegServices As MusterContainer
    Private WithEvents oQS As MUSTER.BusinessLogic.pSearch
    Private MyGUID As New System.Guid
    Private bolErrorOccurred As Boolean = False
    Public Event CompanySearchSelection(ByVal CompanyID As Integer, ByVal LicenseeID As Integer, ByVal AssocID As Integer, ByVal Search_Type As String)
    Public Event CompanySearchResults(ByVal nCount As Integer, ByVal strSrc As String)
    ' Public Event SearchResults(ByVal nCount As Integer, ByVal strSrc As String)
#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef ParentForm As Windows.Forms.Form = Nothing)
        MyBase.New()

        MyGUID = System.Guid.NewGuid

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        If Not ParentForm Is Nothing Then
            Me.MdiParent = ParentForm
        End If
        MusterContainer.AppUser.LogEntry("Company Search", MyGUID.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGUID.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGUID.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
        MusterContainer.AppUser.LogExit(MyGUID.ToString)
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
    Friend WithEvents SearchResult As System.Windows.Forms.GroupBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents pnlSearchContainer As System.Windows.Forms.Panel
    Friend WithEvents pnlSearchHeader As System.Windows.Forms.Panel
    Friend WithEvents ugSearchResult As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblResultCount As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.SearchResult = New System.Windows.Forms.GroupBox
        Me.pnlSearchContainer = New System.Windows.Forms.Panel
        Me.ugSearchResult = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlSearchHeader = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.lblResultCount = New System.Windows.Forms.Label
        Me.SearchResult.SuspendLayout()
        Me.pnlSearchContainer.SuspendLayout()
        CType(Me.ugSearchResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlSearchHeader.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SearchResult
        '
        Me.SearchResult.Controls.Add(Me.pnlSearchContainer)
        Me.SearchResult.Controls.Add(Me.pnlSearchHeader)
        Me.SearchResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SearchResult.Location = New System.Drawing.Point(0, 0)
        Me.SearchResult.Name = "SearchResult"
        Me.SearchResult.Size = New System.Drawing.Size(1028, 430)
        Me.SearchResult.TabIndex = 0
        Me.SearchResult.TabStop = False
        Me.SearchResult.Text = "Company Search Result"
        '
        'pnlSearchContainer
        '
        Me.pnlSearchContainer.Controls.Add(Me.ugSearchResult)
        Me.pnlSearchContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlSearchContainer.Location = New System.Drawing.Point(3, 72)
        Me.pnlSearchContainer.Name = "pnlSearchContainer"
        Me.pnlSearchContainer.Size = New System.Drawing.Size(1022, 355)
        Me.pnlSearchContainer.TabIndex = 1
        '
        'ugSearchResult
        '
        Me.ugSearchResult.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugSearchResult.DisplayLayout.Override.CellAppearance = Appearance1
        Me.ugSearchResult.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance2.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugSearchResult.DisplayLayout.Override.RowAppearance = Appearance2
        Me.ugSearchResult.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugSearchResult.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugSearchResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugSearchResult.Location = New System.Drawing.Point(0, 0)
        Me.ugSearchResult.Name = "ugSearchResult"
        Me.ugSearchResult.Size = New System.Drawing.Size(1022, 355)
        Me.ugSearchResult.TabIndex = 0
        Me.ugSearchResult.Text = "Company Search Result"
        '
        'pnlSearchHeader
        '
        Me.pnlSearchHeader.Controls.Add(Me.Panel1)
        Me.pnlSearchHeader.Controls.Add(Me.lblResultCount)
        Me.pnlSearchHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlSearchHeader.Location = New System.Drawing.Point(3, 16)
        Me.pnlSearchHeader.Name = "pnlSearchHeader"
        Me.pnlSearchHeader.Size = New System.Drawing.Size(1022, 56)
        Me.pnlSearchHeader.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnClose)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1022, 56)
        Me.Panel1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 16)
        Me.Label1.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(608, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 0
        Me.btnClose.Text = "Close"
        '
        'lblResultCount
        '
        Me.lblResultCount.AutoSize = True
        Me.lblResultCount.Location = New System.Drawing.Point(8, 8)
        Me.lblResultCount.Name = "lblResultCount"
        Me.lblResultCount.Size = New System.Drawing.Size(0, 16)
        Me.lblResultCount.TabIndex = 0
        '
        'CompanySearchResults
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1028, 430)
        Me.Controls.Add(Me.SearchResult)
        Me.Name = "CompanySearchResults"
        Me.Text = "Company Search Result"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SearchResult.ResumeLayout(False)
        Me.pnlSearchContainer.ResumeLayout(False)
        CType(Me.ugSearchResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlSearchHeader.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Function NewShowResults(ByRef objQS As MUSTER.BusinessLogic.pSearch)
        Dim dsResults As New DataSet
        Try
            oQS = objQS
            dsResults = oQS.GetResult
            If dsResults Is Nothing Then
                If Not Me.Visible Then Me.Show()
                Me.Close()
                Exit Function
            End If
            If dsResults.Tables.Count > 0 Then
                Dim dtResults As DataTable = dsResults.Tables(0)
                Select Case dtResults.Rows.Count

                    Case 0
                        If Not Me.Visible Then Me.Show()
                        Me.Close()
                        Exit Function
                    Case 1
                        If oQS.Filter = "Company Name" Then
                            RaiseEvent CompanySearchSelection(IIf(IsDBNull(dtResults.Rows(0).Item("CompanyID")), 0, dtResults.Rows(0).Item("CompanyID")), 0, 0, oQS.Filter)
                        Else
                            RaiseEvent CompanySearchSelection(IIf(IsDBNull(dtResults.Rows(0).Item("Company_ID")), 0, dtResults.Rows(0).Item("Company_ID")), IIf(IsDBNull(dtResults.Rows(0).Item("LicenseeID")), 0, dtResults.Rows(0).Item("LicenseeID")), IIf(IsDBNull(dtResults.Rows(0).Item("companyAddress")), 0, dtResults.Rows(0).Item("companyAddress")), oQS.Filter)
                        End If
                        If Not Me.Visible Then
                            Me.Show()
                            Me.Close()
                        End If
                        Exit Select
                    Case Else
                        ugSearchResult.DataSource = dtResults
                        'ugSearchResult.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                        If oQS.Filter.IndexOf("Company Name") > -1 Then
                            'ugSearchResult.DisplayLayout.Bands(0).Columns("CompanyID").Width = 75
                            ugSearchResult.DisplayLayout.Bands(0).Columns("Company Name").Width = 200
                            ugSearchResult.DisplayLayout.Bands(0).Columns("Company Address").Width = 200
                            ugSearchResult.DisplayLayout.Bands(0).Columns("Company Address").VertScrollBar = True
                            ugSearchResult.DisplayLayout.Bands(0).Columns("Company Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                            ugSearchResult.DisplayLayout.Bands(0).Columns("City").Width = 100
                            ugSearchResult.DisplayLayout.Bands(0).Columns("State").Width = 45
                            ugSearchResult.DisplayLayout.Bands(0).Columns("Zip").Width = 45
                            ugSearchResult.DisplayLayout.Bands(0).Columns("Installer/Closures").Width = 130
                            ugSearchResult.DisplayLayout.Bands(0).Columns("Closures").Width = 80
                            ugSearchResult.DisplayLayout.Bands(0).Columns("ERAC").Width = 45
                            ugSearchResult.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
                            ugSearchResult.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
                            ugSearchResult.DisplayLayout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
                            ugSearchResult.DisplayLayout.Bands(0).Override.RowSizingAutoMaxLines = 5
                            ugSearchResult.DisplayLayout.Bands(0).Columns("CompanyID").Hidden = True
                            Me.ugSearchResult.Text = "Company Search Result"
                            Me.Text = "Company Search Result"
                            Me.SearchResult.Text = "Company Search Result"
                        Else
                            If oQS.Filter.IndexOf("Licensee Name") > -1 Then
                                'ugSearchResult.DisplayLayout.Bands(0).Columns("LicenseeID").Width = 75
                                'ugSearchResult.DisplayLayout.Bands(0).Columns("Licensee").Width = 125
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Licensee").Hidden = True

                                ugSearchResult.DisplayLayout.Bands(0).Columns("FirstName").Width = 65
                                ugSearchResult.DisplayLayout.Bands(0).Columns("LastName").Width = 60

                                ugSearchResult.DisplayLayout.Bands(0).Columns("CompanyName").Width = 182
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Address").Width = 150
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Address").VertScrollBar = True
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                                ugSearchResult.DisplayLayout.Bands(0).Columns("City").Width = 80
                                ugSearchResult.DisplayLayout.Bands(0).Columns("State").Width = 40
                                ugSearchResult.DisplayLayout.Bands(0).Columns("LicenseeNo").Width = 65
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Cert_Type").Width = 60
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Status").Width = 90
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Phone_Number_One").Width = 90
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Phone_Number_One").Header.Caption = "Phone1"
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Expire_Date").Width = 70
                                ugSearchResult.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
                                ugSearchResult.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
                                ugSearchResult.DisplayLayout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
                                ugSearchResult.DisplayLayout.Bands(0).Override.RowSizingAutoMaxLines = 5

                                ugSearchResult.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
                                ugSearchResult.DisplayLayout.Bands(0).Columns("Company_id").Hidden = True
                                ugSearchResult.DisplayLayout.Bands(0).Columns("CompanyAddress").Hidden = True

                                ugSearchResult.DisplayLayout.Bands(0).Columns("LastName").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                                Me.ugSearchResult.Text = "Licensee Search Result"
                                Me.Text = "Licensee Search Result"
                                Me.SearchResult.Text = "Licensee Search Result"
                            Else
                                If oQS.Filter.IndexOf("Manager Name") > -1 Then
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Licensee").Hidden = True

                                    ugSearchResult.DisplayLayout.Bands(0).Columns("FirstName").Width = 65
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("LastName").Width = 60

                                    ugSearchResult.DisplayLayout.Bands(0).Columns("CompanyName").Width = 182
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Address").Width = 150
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Address").VertScrollBar = True
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("City").Width = 80
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("State").Width = 40
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("LicenseeNo").Hidden = True
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Cert_Type").Hidden = True
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Status").Width = 90
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Phone_Number_One").Width = 90
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Phone_Number_One").Header.Caption = "Phone1"
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Expire_Date").Width = 70
                                    ugSearchResult.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
                                    ugSearchResult.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
                                    ugSearchResult.DisplayLayout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
                                    ugSearchResult.DisplayLayout.Bands(0).Override.RowSizingAutoMaxLines = 5

                                    ugSearchResult.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("Company_id").Hidden = True
                                    ugSearchResult.DisplayLayout.Bands(0).Columns("CompanyAddress").Hidden = True

                                    ugSearchResult.DisplayLayout.Bands(0).Columns("LastName").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                                    Me.ugSearchResult.Text = "Compliance Manager Search Result"
                                    Me.Text = "Compliance Manager Search Result"
                                    Me.SearchResult.Text = "Compliance Manager Search Result"
                                End If
                            End If
                        End If

                        ugSearchResult.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
                        ugSearchResult.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
                        ugSearchResult.DisplayLayout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
                        ugSearchResult.DisplayLayout.Bands(0).Override.RowSizingAutoMaxLines = 5
                        ugSearchResult.ActiveRow = ugSearchResult.Rows(0)
                        Me.ugSearchResult.DisplayLayout.Bands(0).Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

                        Me.Show()
                        Me.BringToFront()
                End Select
            End If

        Catch ex As Exception
            Throw ex
            If bolErrorOccurred Then
                If Not Me.Visible Then Me.Show()
                Me.Close()
            End If
        End Try
    End Function
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            If Not Me.Visible Then Me.Show()
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SearchResults_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugSearchResult_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugSearchResult.AfterSortChange
        If ugSearchResult.Rows.Count > 0 Then
            ugSearchResult.ActiveRow = ugSearchResult.Rows(0)
        End If
    End Sub
    Private Sub ugSearchResult_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugSearchResult.DoubleClick

        Dim nFacilityId As Integer
        Dim nOwnerId As Integer

        If Not UIUtilsInfragistics.WinGridRowDblClicked(ugSearchResult, New System.EventArgs) Then Exit Sub

        If Me.Cursor Is Cursors.WaitCursor Then
            'Exit Sub
        End If
        Try
            Me.Cursor = Cursors.WaitCursor

            If oQS.Filter = "Company Name" Then
                RaiseEvent CompanySearchSelection(IIf(IsDBNull(ugSearchResult.ActiveRow.Cells("CompanyID").Value), 0, ugSearchResult.ActiveRow.Cells("CompanyID").Value), 0, 0, oQS.Filter)
            Else
                RaiseEvent CompanySearchSelection(IIf(IsDBNull(ugSearchResult.ActiveRow.Cells("Company_ID").Value), 0, ugSearchResult.ActiveRow.Cells("Company_ID").Value), IIf(IsDBNull(ugSearchResult.ActiveRow.Cells("LicenseeID").Value), 0, ugSearchResult.ActiveRow.Cells("LicenseeID").Value), IIf(IsDBNull(ugSearchResult.ActiveRow.Cells("companyAddress").Value), 0, ugSearchResult.ActiveRow.Cells("companyAddress").Value), oQS.Filter)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub SearchResults_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

    End Sub

    Private Sub SearchResults_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

    End Sub
    Private Sub oQS_SearchResultCount(ByVal nCount As Integer, ByVal strSrc As String) Handles oQS.SearchResults
        Me.Label1.Text = "Your Search Resulted in " + nCount.ToString & " records."
        RaiseEvent CompanySearchResults(nCount, strSrc)
    End Sub
End Class
