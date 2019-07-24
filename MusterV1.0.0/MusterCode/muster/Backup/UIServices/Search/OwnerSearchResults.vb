Public Class OwnerSearchResults
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.OwnerSearchResults.vb
    '   Provides the interface for displaying quick search results.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1        JVC2    1/26/05    Changed NewShowResults to check for FacID is null
    '                                   when formulating return to client.
    '  1.2        AN      02/10/05    Integrated AppFlags new object model
    '                                 
    '-------------------------------------------------------------------------------
    '
    'TODO - Remove comment from VSS version 2/9/05 - JVC 2
    '
    Inherits System.Windows.Forms.Form
    Friend WithEvents frmRegServices As MusterContainer
    Private WithEvents oQS As Muster.BusinessLogic.pSearch
    Private MyGUID As System.Guid
    Private bolErrorOccurred As Boolean = False
    Public Event SearchResultSelection(ByVal OwnerID As Integer, ByVal FacilityID As Integer, ByVal Search_Type As String)
    Public Event OwnerSearchResultErr(ByVal MsgStr As String, ByVal strColumnName As String, ByVal strSrc As String)
    Public Event SearchResults(ByVal nCount As Integer, ByVal strSrc As String)


#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef ParentForm As Windows.Forms.Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        If Not ParentForm Is Nothing Then
            Me.MdiParent = ParentForm
        End If

        MusterContainer.AppUser.LogEntry("Quick Search", MyGUID.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGUID.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        '
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
    Friend WithEvents ctxMenuRegistration As System.Windows.Forms.ContextMenu
    Friend WithEvents mnuItemOwner As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemLust As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemOwnerDetail As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemFacilityDetail As System.Windows.Forms.MenuItem
    Friend WithEvents mnuItemGlobalFlag As System.Windows.Forms.MenuItem
    Friend WithEvents grpSearchResult As System.Windows.Forms.GroupBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ugQuickSearch As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.grpSearchResult = New System.Windows.Forms.GroupBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ugQuickSearch = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnClose = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnPrint = New System.Windows.Forms.Button
        Me.ctxMenuRegistration = New System.Windows.Forms.ContextMenu
        Me.mnuItemOwner = New System.Windows.Forms.MenuItem
        Me.mnuItemOwnerDetail = New System.Windows.Forms.MenuItem
        Me.mnuItemFacilityDetail = New System.Windows.Forms.MenuItem
        Me.mnuItemGlobalFlag = New System.Windows.Forms.MenuItem
        Me.mnuItemLust = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.grpSearchResult.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.ugQuickSearch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpSearchResult
        '
        Me.grpSearchResult.Controls.Add(Me.Panel1)
        Me.grpSearchResult.Controls.Add(Me.Panel2)
        Me.grpSearchResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpSearchResult.Location = New System.Drawing.Point(0, 0)
        Me.grpSearchResult.Name = "grpSearchResult"
        Me.grpSearchResult.Size = New System.Drawing.Size(1028, 430)
        Me.grpSearchResult.TabIndex = 0
        Me.grpSearchResult.TabStop = False
        Me.grpSearchResult.Text = "Owner Search Result"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ugQuickSearch)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 56)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1022, 371)
        Me.Panel1.TabIndex = 5
        '
        'ugQuickSearch
        '
        Me.ugQuickSearch.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugQuickSearch.DisplayLayout.Override.CellAppearance = Appearance1
        Me.ugQuickSearch.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance2.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugQuickSearch.DisplayLayout.Override.RowAppearance = Appearance2
        Me.ugQuickSearch.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugQuickSearch.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugQuickSearch.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugQuickSearch.Location = New System.Drawing.Point(0, 0)
        Me.ugQuickSearch.Name = "ugQuickSearch"
        Me.ugQuickSearch.Size = New System.Drawing.Size(1022, 371)
        Me.ugQuickSearch.TabIndex = 2
        Me.ugQuickSearch.Text = "Owner Search Result"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnClose)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.btnPrint)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(3, 16)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1022, 40)
        Me.Panel2.TabIndex = 4
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(608, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 16)
        Me.Label1.TabIndex = 0
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(696, 8)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.TabIndex = 1
        Me.btnPrint.Text = "Print"
        '
        'ctxMenuRegistration
        '
        Me.ctxMenuRegistration.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemOwner, Me.mnuItemLust, Me.MenuItem6})
        '
        'mnuItemOwner
        '
        Me.mnuItemOwner.Index = 0
        Me.mnuItemOwner.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuItemOwnerDetail, Me.mnuItemFacilityDetail, Me.mnuItemGlobalFlag})
        Me.mnuItemOwner.Text = "Owner"
        Me.mnuItemOwner.Visible = False
        '
        'mnuItemOwnerDetail
        '
        Me.mnuItemOwnerDetail.Index = 0
        Me.mnuItemOwnerDetail.Text = "Show Owner Details"
        '
        'mnuItemFacilityDetail
        '
        Me.mnuItemFacilityDetail.Index = 1
        Me.mnuItemFacilityDetail.Text = "Show Facility Details"
        '
        'mnuItemGlobalFlag
        '
        Me.mnuItemGlobalFlag.Index = 2
        Me.mnuItemGlobalFlag.Text = "Show Global Flags"
        '
        'mnuItemLust
        '
        Me.mnuItemLust.Index = 1
        Me.mnuItemLust.Text = "LUST"
        Me.mnuItemLust.Visible = False
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 2
        Me.MenuItem6.Text = "Closure"
        Me.MenuItem6.Visible = False
        '
        'OwnerSearchResults
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(1028, 430)
        Me.ContextMenu = Me.ctxMenuRegistration
        Me.Controls.Add(Me.grpSearchResult)
        Me.Name = "OwnerSearchResults"
        Me.Text = "Owner Search Result"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grpSearchResult.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        CType(Me.ugQuickSearch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public Function NewShowResults(ByRef objQS As Muster.BusinessLogic.pSearch)
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
                        RaiseEvent SearchResultSelection(0, 0, Nothing)
                        Me.Dispose()
                    Case 1
                        RaiseEvent SearchResultSelection(dtResults.Rows(0).Item("OWNERID"), IIf(IsDBNull(dtResults.Rows(0).Item("FACILITYID")), 0, dtResults.Rows(0).Item("FACILITYID")), oQS.Filter)
                        If Not Me.Visible Then
                            Me.Show()
                            Me.Close()
                        End If
                        Exit Select
                    Case Else
                        ugQuickSearch.DataSource = dtResults
                        ugQuickSearch.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
                        ugQuickSearch.DisplayLayout.Bands(0).AutoPreviewEnabled = True
                        ugQuickSearch.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                        ugQuickSearch.DisplayLayout.Bands(0).Columns("OwnerID").Width = 75
                        ugQuickSearch.DisplayLayout.Bands(0).Columns("OwnerName").Width = 200
                        ugQuickSearch.DisplayLayout.Bands(0).Columns("Address").Width = 250
                        ugQuickSearch.DisplayLayout.Bands(0).Columns("Address").VertScrollBar = True
                        ugQuickSearch.DisplayLayout.Bands(0).Columns("Address").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
                        ugQuickSearch.DisplayLayout.Bands(0).Columns("City").Width = 100
                        ugQuickSearch.DisplayLayout.Bands(0).Columns("State").Width = 50
                        ugQuickSearch.DisplayLayout.Bands(0).Columns("Zip").Width = 75
                        ugQuickSearch.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
                        ugQuickSearch.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
                        ugQuickSearch.DisplayLayout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
                        ugQuickSearch.DisplayLayout.Bands(0).Override.RowSizingAutoMaxLines = 5
                        ugQuickSearch.ActiveRow = ugQuickSearch.Rows(0)
                        Me.ugQuickSearch.DisplayLayout.Bands(0).Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                        If oQS.Filter.IndexOf("Facility") > -1 Then
                            ugQuickSearch.DisplayLayout.Bands(0).Columns("FacilityID").Width = 75
                            ugQuickSearch.DisplayLayout.Bands(0).Columns("FacilityName").Width = 200
                            ugQuickSearch.DisplayLayout.Bands(0).Columns("FacilityID").Hidden = False
                            ugQuickSearch.DisplayLayout.Bands(0).Columns("FacilityName").Hidden = False
                            Me.ugQuickSearch.Text = "Facility Search Result"
                            Me.Text = "Facility Search Result"
                            Me.grpSearchResult.Text = "Facility Search Result"
                        Else
                            Me.ugQuickSearch.Text = "Owner Search Result"
                            Me.Text = "Owner Search Result"
                            Me.grpSearchResult.Text = "Owner Search Result"
                            ugQuickSearch.DisplayLayout.Bands(0).Columns("FacilityID").Hidden = True
                            ugQuickSearch.DisplayLayout.Bands(0).Columns("FacilityName").Hidden = True
                        End If
                        Me.Show()
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
    Private Sub OwnerSearchResults_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugQuickSearch_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugQuickSearch.DoubleClick

        Dim nFacilityId As Integer
        Dim nOwnerId As Integer

        If Not UIUtilsInfragistics.WinGridRowDblClicked(ugQuickSearch, New System.EventArgs) Then Exit Sub

        If Me.Cursor Is Cursors.WaitCursor Then
            'Exit Sub
        End If
        Try
            Me.Cursor = Cursors.WaitCursor

            RaiseEvent SearchResultSelection(IIf(IsDBNull(ugQuickSearch.ActiveRow.Cells("OWNERID").Value), 0, ugQuickSearch.ActiveRow.Cells("OWNERID").Value), (IIf(IsDBNull(ugQuickSearch.ActiveRow.Cells("FACILITYID").Value), 0, ugQuickSearch.ActiveRow.Cells("FACILITYID").Value)), oQS.Filter)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub ugQuickSearch_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugQuickSearch.AfterSortChange
        If ugQuickSearch.Rows.Count > 0 Then
            ugQuickSearch.ActiveRow = ugQuickSearch.Rows(0)
        End If

    End Sub
    Private Sub oQS_SearchErr(ByVal MsgStr As String, ByVal strColumnName As String, ByVal strSrc As String) Handles oQS.SearchErr
        RaiseEvent OwnerSearchResultErr(MsgStr, strColumnName, strSrc)
        ' MsgBox(MsgStr & vbCrLf & "Column : " & strColumnName & vbCrLf & "Source : " & strSrc)
        bolErrorOccurred = True
    End Sub
    Private Sub oQS_SearchResultCount(ByVal nCount As Integer, ByVal strSrc As String) Handles oQS.SearchResults
        Me.Label1.Text = "Your Search Resulted in " + nCount.ToString & " records."
        RaiseEvent SearchResults(nCount, strSrc)
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        ugQuickSearch.PrintPreview(Infragistics.Win.UltraWinGrid.RowPropertyCategories.All)
    End Sub

    Private Sub ugQuickSearch_InitializePrintPreview(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelablePrintPreviewEventArgs) Handles ugQuickSearch.InitializePrintPreview
        UIUtilsGen.ug_InitializePrintPreview(sender, e, oQS.Filter + " search: " + oQS.Keyword, _
                                                Infragistics.Win.UltraWinGrid.ColumnClipMode.RepeatClippedColumns, _
                                                True, Printing.PrintRange.AllPages, _
                                                True, , Infragistics.Win.DefaultableBoolean.True, , 1)
    End Sub
End Class
