Public Class DocumentViewControl
    Inherits System.Windows.Forms.UserControl

    'Private WithEvents WordApp As Word.Application
    Dim result As DialogResult
    Dim rp As New Remove_Pencil
    Dim bolLoading As Boolean = False
    Dim dsDocuments As DataSet
    Public EntityID As Integer = 0
    Public ModuleID As Integer = 0
    Public EntityType As Integer = 0

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
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
    Friend WithEvents cmbYear As System.Windows.Forms.ComboBox
    Friend WithEvents cmbModule As System.Windows.Forms.ComboBox
    Friend WithEvents lblModule As System.Windows.Forms.Label
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents pnlMain As System.Windows.Forms.Panel
    Friend WithEvents pnlDocumentsTop As System.Windows.Forms.Panel
    Friend WithEvents chkShowAllUserDocuments As System.Windows.Forms.CheckBox
    Friend WithEvents pnlDocumentsMain As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ugDocuments As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbYear = New System.Windows.Forms.ComboBox
        Me.cmbModule = New System.Windows.Forms.ComboBox
        Me.lblModule = New System.Windows.Forms.Label
        Me.lblYear = New System.Windows.Forms.Label
        Me.pnlMain = New System.Windows.Forms.Panel
        Me.pnlDocumentsMain = New System.Windows.Forms.Panel
        Me.ugDocuments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlDocumentsTop = New System.Windows.Forms.Panel
        Me.chkShowAllUserDocuments = New System.Windows.Forms.CheckBox
        Me.pnlMain.SuspendLayout()
        Me.pnlDocumentsMain.SuspendLayout()
        CType(Me.ugDocuments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlDocumentsTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbYear
        '
        Me.cmbYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbYear.DropDownWidth = 60
        Me.cmbYear.ItemHeight = 13
        Me.cmbYear.Location = New System.Drawing.Point(296, 8)
        Me.cmbYear.Name = "cmbYear"
        Me.cmbYear.Size = New System.Drawing.Size(56, 21)
        Me.cmbYear.TabIndex = 142
        '
        'cmbModule
        '
        Me.cmbModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbModule.DropDownWidth = 160
        Me.cmbModule.ItemHeight = 13
        Me.cmbModule.Location = New System.Drawing.Point(88, 8)
        Me.cmbModule.Name = "cmbModule"
        Me.cmbModule.Size = New System.Drawing.Size(144, 21)
        Me.cmbModule.TabIndex = 140
        '
        'lblModule
        '
        Me.lblModule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModule.Location = New System.Drawing.Point(24, 8)
        Me.lblModule.Name = "lblModule"
        Me.lblModule.Size = New System.Drawing.Size(48, 23)
        Me.lblModule.TabIndex = 141
        Me.lblModule.Text = "Module"
        '
        'lblYear
        '
        Me.lblYear.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear.Location = New System.Drawing.Point(248, 8)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(40, 23)
        Me.lblYear.TabIndex = 143
        Me.lblYear.Text = "Year"
        '
        'pnlMain
        '
        Me.pnlMain.Controls.Add(Me.pnlDocumentsMain)
        Me.pnlMain.Controls.Add(Me.pnlDocumentsTop)
        Me.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMain.Location = New System.Drawing.Point(0, 0)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(832, 280)
        Me.pnlMain.TabIndex = 146
        '
        'pnlDocumentsMain
        '
        Me.pnlDocumentsMain.Controls.Add(Me.ugDocuments)
        Me.pnlDocumentsMain.Controls.Add(Me.Label1)
        Me.pnlDocumentsMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlDocumentsMain.Location = New System.Drawing.Point(0, 40)
        Me.pnlDocumentsMain.Name = "pnlDocumentsMain"
        Me.pnlDocumentsMain.Size = New System.Drawing.Size(832, 240)
        Me.pnlDocumentsMain.TabIndex = 149
        '
        'ugDocuments
        '
        Me.ugDocuments.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugDocuments.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugDocuments.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugDocuments.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugDocuments.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugDocuments.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugDocuments.Location = New System.Drawing.Point(0, 0)
        Me.ugDocuments.Name = "ugDocuments"
        Me.ugDocuments.Size = New System.Drawing.Size(832, 240)
        Me.ugDocuments.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(792, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(7, 23)
        Me.Label1.TabIndex = 2
        '
        'pnlDocumentsTop
        '
        Me.pnlDocumentsTop.Controls.Add(Me.chkShowAllUserDocuments)
        Me.pnlDocumentsTop.Controls.Add(Me.lblYear)
        Me.pnlDocumentsTop.Controls.Add(Me.lblModule)
        Me.pnlDocumentsTop.Controls.Add(Me.cmbModule)
        Me.pnlDocumentsTop.Controls.Add(Me.cmbYear)
        Me.pnlDocumentsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlDocumentsTop.DockPadding.All = 3
        Me.pnlDocumentsTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlDocumentsTop.Name = "pnlDocumentsTop"
        Me.pnlDocumentsTop.Size = New System.Drawing.Size(832, 40)
        Me.pnlDocumentsTop.TabIndex = 148
        '
        'chkShowAllUserDocuments
        '
        Me.chkShowAllUserDocuments.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowAllUserDocuments.Location = New System.Drawing.Point(376, 8)
        Me.chkShowAllUserDocuments.Name = "chkShowAllUserDocuments"
        Me.chkShowAllUserDocuments.Size = New System.Drawing.Size(112, 16)
        Me.chkShowAllUserDocuments.TabIndex = 2
        Me.chkShowAllUserDocuments.Tag = "646"
        Me.chkShowAllUserDocuments.Text = "Show All Users"
        Me.chkShowAllUserDocuments.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'DocumentViewControl
        '
        Me.AutoScroll = True
        Me.Controls.Add(Me.pnlMain)
        Me.Name = "DocumentViewControl"
        Me.Size = New System.Drawing.Size(832, 280)
        Me.pnlMain.ResumeLayout(False)
        Me.pnlDocumentsMain.ResumeLayout(False)
        CType(Me.ugDocuments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlDocumentsTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub PopulateLetters(Optional ByVal retrieveRecords As Boolean = False)
        Try
            Dim dtTable As DataTable
            Dim drow As DataRow

            If retrieveRecords Then
                dsDocuments = MusterContainer.pLetter.GetManualAndSystemDocuments(IIf(chkShowAllUserDocuments.Checked, String.Empty, MusterContainer.AppUser.ID))
            End If

            dtTable = dsDocuments.Tables(0)
            ugDocuments.DataSource = Nothing
            dtTable.DefaultView.Sort = "[Date Created] ASC"
            'Filter based on Module,Year,ShowAllUsers & Entity Type etc.,

            Select Case EntityType
                Case UIUtilsGen.EntityTypes.Facility
                    dtTable.DefaultView.RowFilter = "Facility_ID = " + EntityID.ToString + _
                          " AND (Module_ID = " + IIf(cmbModule.Text = "" Or cmbModule.Text = "ALL", "Module_ID", UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString) + " )" + _
                          " AND YEAR = " + IIf(cmbYear.Text = "" Or cmbYear.Text = "ALL", "YEAR", UIUtilsGen.GetComboBoxValueString(cmbYear).ToString)
                    '  " AND (Module_ID = " + IIf(cmbModule.Text = "" Or cmbModule.Text = "ALL", "Module_ID", UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString) + " OR TYPE = 'Manual')" + _
                Case UIUtilsGen.EntityTypes.Owner
                    dtTable.DefaultView.RowFilter = "Owner_ID = " + EntityID.ToString + _
                        " AND (Module_ID = " + IIf(cmbModule.Text = "" Or cmbModule.Text = "ALL", "Module_ID", UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString) + " )" + _
                        " AND YEAR = " + IIf(cmbYear.Text = "" Or cmbYear.Text = "ALL", "YEAR", UIUtilsGen.GetComboBoxValueString(cmbYear).ToString)
                Case Else
                    dtTable.DefaultView.RowFilter = "[Entity ID] = " + EntityID.ToString + _
                        " AND (Module_ID = " + IIf(cmbModule.Text = "" Or cmbModule.Text = "ALL", "Module_ID", UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString) + " )" + _
                        " AND YEAR = " + IIf(cmbYear.Text = "" Or cmbYear.Text = "ALL", "YEAR", UIUtilsGen.GetComboBoxValueString(cmbYear).ToString)
            End Select

            'If Not cmbModule.Text = "ALL" Then
            '    If Not cmbModule.SelectedIndex = -1 And Not cmbYear.SelectedIndex = -1 Then
            '        Select Case EntityType
            '            Case 6
            '                dtTable.DefaultView.RowFilter = "([Facility_ID] = " + EntityID.ToString + ") AND (Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString + " OR TYPE = 'Manual') AND YEAR = " + UIUtilsGen.GetComboBoxValueString(cmbYear).ToString
            '            Case 9
            '                dtTable.DefaultView.RowFilter = "([Owner_ID] = " + EntityID.ToString + ") AND (Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString + " OR TYPE = 'Manual') AND YEAR = " + UIUtilsGen.GetComboBoxValueString(cmbYear).ToString
            '            Case 0
            '                dtTable.DefaultView.RowFilter = "([Entity ID] = " + EntityID.ToString + ") AND (Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString + " OR TYPE = 'Manual') AND YEAR = " + UIUtilsGen.GetComboBoxValueString(cmbYear).ToString
            '        End Select
            '    ElseIf Not cmbModule.SelectedIndex = -1 Then
            '        Select Case EntityType
            '            Case 6
            '                dtTable.DefaultView.RowFilter = "([Facility_ID] = " + EntityID.ToString + ") AND Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString
            '            Case 9
            '                dtTable.DefaultView.RowFilter = "([Owner_ID] = " + EntityID.ToString + ") AND Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString
            '            Case 0
            '                dtTable.DefaultView.RowFilter = "([Entity ID] = " + EntityID.ToString + ") AND Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString
            '        End Select
            '    End If
            'Else
            '    If Not cmbModule.SelectedIndex = -1 And Not cmbYear.SelectedIndex = -1 Then
            '        Select Case EntityType
            '            Case 6
            '                dtTable.DefaultView.RowFilter = "([Facility_ID] = " + EntityID.ToString + ") AND YEAR = " + UIUtilsGen.GetComboBoxValueString(cmbYear).ToString
            '            Case 9
            '                dtTable.DefaultView.RowFilter = "([Owner_ID] = " + EntityID.ToString + ") AND YEAR = " + UIUtilsGen.GetComboBoxValueString(cmbYear).ToString
            '            Case 0
            '                dtTable.DefaultView.RowFilter = "([Entity ID] = " + EntityID.ToString + ") AND YEAR = " + UIUtilsGen.GetComboBoxValueString(cmbYear).ToString
            '        End Select
            '    ElseIf Not cmbModule.SelectedIndex = -1 Then
            '        Select Case EntityType
            '            Case 6
            '                dtTable.DefaultView.RowFilter = "([Facility_ID] = " + EntityID.ToString + ")"
            '            Case 9
            '                dtTable.DefaultView.RowFilter = "([Owner_ID] = " + EntityID.ToString + ")"
            '            Case 0
            '                dtTable.DefaultView.RowFilter = "([Entity ID] = " + EntityID.ToString + ")"
            '        End Select
            '    End If
            'End If

            LoadDocumentGrid(dtTable)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub LoadPrimaryModules()
        Try
            bolLoading = True
            Dim dtTable As DataTable = MusterContainer.AppUser.ListModulesUserCanSearch(MusterContainer.AppUser.UserKey)
            Dim drow As DataRow
            drow = dtTable.NewRow
            drow("PROPERTY_NAME") = "ALL"
            drow("PROPERTY_ID") = "0"
            dtTable.Rows.Add(drow)
            dtTable.DefaultView.Sort = "PROPERTY_NAME"
            'dtTable.DefaultView.RowFilter = "PROPERTY_ID NOT IN (894,1303,1311,1312)"
            cmbModule.DataSource = dtTable.DefaultView
            cmbModule.DisplayMember = "PROPERTY_NAME"
            cmbModule.ValueMember = "PROPERTY_ID"
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub LoadCalendarYear()
        Try
            bolLoading = True
            Dim dsYear As DataSet
            dsYear = MusterContainer.pLetter.GetCalendarYear(IIf(chkShowAllUserDocuments.Checked, String.Empty, MusterContainer.AppUser.ID), 2)
            If dsYear.Tables.Count <= 0 Then Exit Sub
            Dim dr As DataRow
            dr = dsYear.Tables(0).NewRow
            dr("DATE_CREATED") = "ALL"
            dsYear.Tables(0).Rows.InsertAt(dr, 0)
            cmbYear.DataSource = dsYear.Tables(0).DefaultView
            cmbYear.DisplayMember = "DATE_CREATED"
            cmbYear.ValueMember = "DATE_CREATED"
            UIUtilsGen.SetComboboxItemByText(cmbYear, Year(Today).ToString)
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub LoadDocumentGrid(ByVal dtTable As DataTable)
        ugDocuments.DataSource = dtTable
        ugDocuments.DrawFilter = rp
    End Sub

    Private Sub cmbModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbModule.SelectedIndexChanged
        Try
            If bolLoading Then Exit Sub
            PopulateLetters(True)

            'If dsDocuments.Tables.Count <= 0 Then Exit Sub
            'ugDocuments.DataSource = Nothing

            'If Not cmbModule.Text = "ALL" Then
            '    Select Case EntityType
            '        Case 6
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Facility_ID] = " + EntityID.ToString + ") AND (Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString + " OR Type = 'Manual')"
            '        Case 9
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Owner_ID] = " + EntityID.ToString + ") AND (Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString + " OR Type = 'Manual')"
            '        Case 0
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Entity ID] = " + EntityID.ToString + ") AND (Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString + " OR Type = 'Manual')"
            '    End Select
            'Else
            '    Select Case EntityType
            '        Case 6
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Facility_ID] = " + EntityID.ToString + ")"
            '        Case 9
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Owner_ID] = " + EntityID.ToString + ")"
            '        Case 0
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Entity ID] = " + EntityID.ToString + ")"
            '    End Select

            'End If
            'LoadDocumentGrid(dsDocuments.Tables(0))
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbYear.SelectedIndexChanged
        Try
            If bolLoading Then Exit Sub
            PopulateLetters(False)

            'If dsDocuments.Tables.Count <= 0 Then Exit Sub
            'ugDocuments.DataSource = Nothing

            'If Not cmbModule.Text = "ALL" Then
            '    Select Case EntityType
            '        Case 6
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Facility_ID] = " + EntityID.ToString + " OR [Facility_ID] = 0) AND (Module_ID = " + IIf(cmbModule.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString, "Module_ID") + " OR Type = 'Manual') " + " AND YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, "YEAR")
            '        Case 9
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Owner_ID] = " + EntityID.ToString + " OR [Owner_ID] = 0) AND (Module_ID = " + IIf(cmbModule.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString, "Module_ID") + " OR Type = 'Manual') " + " AND YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, "YEAR")
            '        Case 0
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Entity ID] = " + EntityID.ToString + " OR [Entity ID] = 0) AND (Module_ID = " + IIf(cmbModule.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString, "Module_ID") + " OR Type = 'Manual') " + " AND YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, "YEAR")
            '    End Select
            'Else
            '    Select Case EntityType
            '        Case 6
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Facility_ID] = " + EntityID.ToString + " OR [Facility_ID] = 0)" + " AND YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, "YEAR")
            '        Case 9
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Owner_ID] = " + EntityID.ToString + " OR [Owner_ID] = 0)" + " AND YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, "YEAR")
            '        Case 0
            '            dsDocuments.Tables(0).DefaultView.RowFilter = "([Entity ID] = " + EntityID.ToString + " OR [Entity ID] = 0)" + " AND YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, "YEAR")
            '    End Select

            'End If
            'LoadDocumentGrid(dsDocuments.Tables(0))
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub chkShowAllUserDocuments_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowAllUserDocuments.CheckedChanged
        Try
            PopulateLetters(True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub LoadDocumentsGrid(ByVal nEntityID As Integer, Optional ByVal nEntityType As Integer = 0, Optional ByVal nModuleID As Integer = 0)
        Try
            EntityID = nEntityID
            EntityType = nEntityType
            ModuleID = nModuleID

            LoadPrimaryModules()
            LoadCalendarYear()
            bolLoading = True
            If ModuleID <> 0 Then
                cmbModule.SelectedValue = ModuleID
            Else
                cmbModule.SelectedValue = MusterContainer.AppUser.DefaultModule
            End If

            PopulateLetters(True)
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugDocuments_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugDocuments.DoubleClick
        Dim SrcDoc As Word.Document
        Try

            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            'If Not UIUtilsInfragistics.WinGridRowDblClicked(ugDocuments, New System.EventArgs) Then Exit Sub

            Dim strPath As String = String.Empty

            strPath = Trim(ugDocuments.ActiveRow.Cells("Document Location").Text.ToString)

            If Not strPath.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()) Then
                strPath &= System.IO.Path.DirectorySeparatorChar
            End If

            strPath = strPath + ugDocuments.ActiveRow.Cells("DOCUMENT NAME").Text.ToString

            If Not System.IO.File.Exists(strPath) Then
                MessageBox.Show("Invalid path. This document might have been manually archived. " + vbCrLf + strPath)
                Exit Sub
            End If

            If strPath.ToUpper.IndexOf(".PDF") > -1 Then
                UIUtilsGen.OpenInPDFFile(strPath)
            Else
                '--
                If System.IO.File.Exists(strPath + "x") Then
                    System.Diagnostics.Process.Start(strPath + "x")
                Else
                    System.Diagnostics.Process.Start(strPath)
                End If
            End If
            'WordApp = UIUtilsGen.GetWordApp
            'WordApp.Visible = True

            'If Not ugDocuments.Rows.Count <= 0 Then
            'SrcDoc = WordApp.Documents.Open(strPath)
            'End If
        Catch ex As Exception
            'UIUtilsGen.Delay(, 2)
            ' cannot quit word. if there are other open docs, it will automatically close those docs
            'If Not WordApp Is Nothing Then
            '    WordApp.Quit(False)
            'End If
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot open the file: " + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            'UIUtilsGen.Delay(, 2)
            SrcDoc = Nothing
            'WordApp = Nothing
            Me.ugDocuments.Focus()
            If ugDocuments.Rows.Count > 0 Then
                If ugDocuments.ActiveRow Is Nothing Then
                    ugDocuments.ActiveRow = ugDocuments.Rows(0)
                End If
            End If
        End Try
    End Sub
    Private Sub ugDocuments_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugDocuments.AfterCellUpdate
        If bolLoading Then Exit Sub
        Dim isManualDoc As Boolean = False
        Try
            If e.Cell.Row.Cells("Type").Text.ToUpper = "SYSTEM" Then
                isManualDoc = False
            Else
                isManualDoc = True
            End If
            MusterContainer.pLetter.SaveDocDescription(e.Cell.Row.Cells("document_id").Value, e.Cell.Text, isManualDoc, MusterContainer.AppUser.UserKey)
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot modify description of the file: " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugDocuments_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugDocuments.InitializeLayout
        Try
            e.Layout.Bands(0).Columns("Year").Width = 50
            e.Layout.Bands(0).Columns("Module").Width = 100
            e.Layout.Bands(0).Columns("Document Name").Width = 300
            e.Layout.Bands(0).Columns("Description").Width = 300
            e.Layout.Bands(0).Columns("EventID").Width = 60
            e.Layout.Bands(0).Columns("User").Width = 100
            e.Layout.Bands(0).Columns("Date Created").Width = 75
            e.Layout.Bands(0).Columns("Type").Width = 65

            e.Layout.Bands(0).Columns("Entity Type").Hidden = True
            e.Layout.Bands(0).Columns("Entity ID").Hidden = True
            e.Layout.Bands(0).Columns("Document Type").Hidden = True
            e.Layout.Bands(0).Columns("Date Printed").Hidden = True
            e.Layout.Bands(0).Columns("Document Location").Hidden = True
            e.Layout.Bands(0).Columns("Entity Type ID").Hidden = True
            e.Layout.Bands(0).Columns("Created_By").Hidden = True
            e.Layout.Bands(0).Columns("last_edited_by").Hidden = True
            e.Layout.Bands(0).Columns("date_last_edited").Hidden = True
            e.Layout.Bands(0).Columns("MODULE_ID").Hidden = True
            'e.Layout.Bands(0).Columns("DOCUMENT_ID").Hidden = True
            e.Layout.Bands(0).Columns("Facility_ID").Hidden = True
            e.Layout.Bands(0).Columns("Owner_ID").Hidden = True
            e.Layout.Bands(0).Columns("EVENT_ID").Hidden = True
            e.Layout.Bands(0).Columns("EventID").Hidden = True
            e.Layout.Bands(0).Columns("DOCUMENT_ID").Hidden = True

            e.Layout.Bands(0).Columns("Description").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(0).Columns("Year").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("Module").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("Document Name").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("EventID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("User").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("Date Created").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("Type").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("EVENT_SEQUENCE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("EVENT_TYPE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(0).Columns("EVENT_SEQUENCE").Header.Caption = "Event #"
            e.Layout.Bands(0).Columns("EVENT_TYPE").Header.Caption = "Event Type"

            e.Layout.Bands(0).Columns("Description").FieldLen = 500

            e.Layout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            e.Layout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
            e.Layout.Override.RowSizingAutoMaxLines = 5
            'e.Layout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
            'e.Layout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
            'e.Layout.Bands(0).Override.RowSizingAutoMaxLines = 5
        Catch ex As Exception

        End Try
    End Sub
End Class

