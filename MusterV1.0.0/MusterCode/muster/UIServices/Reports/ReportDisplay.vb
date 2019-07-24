Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.Configuration
Imports Microsoft.Win32.Registry



Public Class ReportDisplay
    Inherits System.Windows.Forms.Form
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.ReportDisplay
    '   Provides the interface for generating reports within the application
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1        JVC2    2/28/05    Added GetFavorites functionality.
    '-------------------------------------------------------------------------------
    '

#Region "User Defined Variables"

    Public WithEvents oRpt As ReportDocument
    Public WithEvents oParameterForm As New CrystalReportsParamForm
    Public oReports As New MUSTER.BusinessLogic.pReport

    Private bolLoading As Boolean = False
    Private dtReports As DataTable
    Private drReports As DataRow
    Private returnVal As String = String.Empty

    Protected REPORT_PATH As String = String.Format("{0}\", MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Reports).ProfileValue)
    Protected TEMPLETTER_PATH As String = String.Format("{0}\", MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue)
    Protected DOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue & "\"

    'Added for closing parama duplcate show on 21st Jan 2015
    Public Shared FlagParamDialogClosed As Boolean = False
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        bolLoading = True
        'Add any initialization after the InitializeComponent() call

        '' Declare variables needed to pass the parameters
        '' to the viewer control.
        'Dim paramFields As New ParameterFields
        'Dim paramField As New ParameterField
        'Dim discreteVal As New ParameterDiscreteValue
        'Dim rangeVal As New ParameterRangeValue


        '' The first parameter is a discrete parameter with multiple values.

        '' Set the name of the parameter field, this must match a 
        '' parameter in the report.
        'paramField.ParameterFieldName = "ZIP Code"

        '' Set the first discrete value and pass it to the parameter
        'discreteVal.Value = "39157"
        'paramField.CurrentValues.Add(discreteVal)

        '' Set the second discrete value and pass it to the parameter.
        '' The discreteVal variable is set to new so the previous settings
        '' will not be overwritten.
        ''discreteVal = New ParameterDiscreteValue
        ''discreteVal.Value = "Aruba Sport"
        ''paramField.CurrentValues.Add(discreteVal)

        '' Add the parameter to the parameter fields collection.
        'paramFields.Add(paramField)

        '' The second parameter is a range value. The paramField variable
        '' is set to new so the previous settings will not be overwritten.
        ''paramField = New ParameterField

        '' Set the name of the parameter field, this must match a
        '' parameter in the report.
        ''paramField.ParameterFieldName = "ZIP Code"

        '' Set the start and end values of the range and pass it to the 'parameter.
        ''rangeVal.StartValue = 42
        ''rangeVal.EndValue = 72
        ''paramField.CurrentValues.Add(rangeVal)

        '' Add the second parameter to the parameter fields collection.
        ''paramFields.Add(paramField)

        '' Set the parameter fields collection into the viewer control.
        'CrystalReportViewer1.ParameterFieldInfo = paramFields


        'oRpt = New ReportDocument
        'oRpt.Load("D:\\Chandra\\Report1.rpt")
        'CrystalReportViewer1.ReportSource = oRpt
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
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents btnDeleteFav As System.Windows.Forms.Button
    Friend WithEvents btnSaveasFavorite As System.Windows.Forms.Button
    Friend WithEvents lblModule As System.Windows.Forms.Label
    Friend WithEvents cboModule As System.Windows.Forms.ComboBox
    Friend WithEvents lblReport As System.Windows.Forms.Label
    Friend WithEvents cboReports As System.Windows.Forms.ComboBox
    Friend WithEvents btnGo As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnDeleteFav = New System.Windows.Forms.Button()
        Me.btnSaveasFavorite = New System.Windows.Forms.Button()
        Me.lblModule = New System.Windows.Forms.Label()
        Me.cboModule = New System.Windows.Forms.ComboBox()
        Me.lblReport = New System.Windows.Forms.Label()
        Me.cboReports = New System.Windows.Forms.ComboBox()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CrystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default
        Me.CrystalReportViewer1.DisplayBackgroundEdge = False
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(8, 56)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ShowGroupTreeButton = False
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(970, 592)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnDeleteFav)
        Me.Panel1.Controls.Add(Me.btnSaveasFavorite)
        Me.Panel1.Controls.Add(Me.lblModule)
        Me.Panel1.Controls.Add(Me.cboModule)
        Me.Panel1.Controls.Add(Me.lblReport)
        Me.Panel1.Controls.Add(Me.cboReports)
        Me.Panel1.Controls.Add(Me.btnGo)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(960, 56)
        Me.Panel1.TabIndex = 13
        '
        'btnDeleteFav
        '
        Me.btnDeleteFav.Location = New System.Drawing.Point(586, 24)
        Me.btnDeleteFav.Name = "btnDeleteFav"
        Me.btnDeleteFav.Size = New System.Drawing.Size(95, 24)
        Me.btnDeleteFav.TabIndex = 19
        Me.btnDeleteFav.Text = "Delete Favorite"
        '
        'btnSaveasFavorite
        '
        Me.btnSaveasFavorite.Location = New System.Drawing.Point(473, 24)
        Me.btnSaveasFavorite.Name = "btnSaveasFavorite"
        Me.btnSaveasFavorite.Size = New System.Drawing.Size(108, 24)
        Me.btnSaveasFavorite.TabIndex = 18
        Me.btnSaveasFavorite.Text = "Save As Favorite"
        '
        'lblModule
        '
        Me.lblModule.Location = New System.Drawing.Point(57, 9)
        Me.lblModule.Name = "lblModule"
        Me.lblModule.Size = New System.Drawing.Size(56, 16)
        Me.lblModule.TabIndex = 17
        Me.lblModule.Text = "Module"
        '
        'cboModule
        '
        Me.cboModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboModule.Location = New System.Drawing.Point(55, 27)
        Me.cboModule.Name = "cboModule"
        Me.cboModule.Size = New System.Drawing.Size(146, 21)
        Me.cboModule.TabIndex = 16
        '
        'lblReport
        '
        Me.lblReport.Location = New System.Drawing.Point(209, 11)
        Me.lblReport.Name = "lblReport"
        Me.lblReport.Size = New System.Drawing.Size(48, 16)
        Me.lblReport.TabIndex = 15
        Me.lblReport.Text = "Reports"
        '
        'cboReports
        '
        Me.cboReports.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReports.Location = New System.Drawing.Point(209, 27)
        Me.cboReports.Name = "cboReports"
        Me.cboReports.Size = New System.Drawing.Size(216, 21)
        Me.cboReports.TabIndex = 14
        '
        'btnGo
        '
        Me.btnGo.Enabled = False
        Me.btnGo.Location = New System.Drawing.Point(427, 25)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(38, 23)
        Me.btnGo.TabIndex = 13
        Me.btnGo.Text = "GO"
        '
        'ReportDisplay
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(968, 526)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CrystalReportViewer1)
        Me.Name = "ReportDisplay"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Report Display"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Old Code"

    '  Private Sub loadModules()
    '     Try
    'Dim dt As DataTable = MusterContainer.AppUser.ListModulesUserHasAccessTo(MusterContainer.AppUser.UserKey)


    'Dim oModProp As New MUSTER.BusinessLogic.pPropertyType("Modules")
    'Dim dtModules As DataTable
    'Dim drModule As DataRow
    'Dim i As Integer = 0

    'dtModules = oModProp.PropertiesTable

    'drModule = dtModules.NewRow
    'drModule.Item("Property Name") = "Favorite Reports"
    'drModule.Item("Property ID") = -1
    'drModule.Item("PropType_ID") = 0
    'dtModules.Rows.InsertAt(drModule, 0)

    'drModule = dtModules.NewRow
    'drModule.Item("Property Name") = "All Reports"
    'drModule.Item("Property ID") = 0
    'drModule.Item("PropType_ID") = 0
    'dtModules.Rows.InsertAt(drModule, dtModules.Rows.Count + 1)

    'dtModules.DefaultView.RowFilter = "PropType_ID=0 OR PropType_ID=" + oModProp.ID.ToString
    'Me.cboModule.DataSource = dtModules.DefaultView
    'Me.cboModule.DisplayMember = "Property Name"
    'Me.cboModule.ValueMember = "Property ID"

    '    Catch ex As Exception
    '       Throw ex
    '  End Try

    ' End Sub


#End Region

#Region "UI Support Routines"



    Private Sub LoadReportFromTickler()

        With _container

            If .GotoReport.Length > 0 Then

                cboModule.Text = .GotoModule
                cboReports.Text = .GotoReport

                .GotoReport = String.Empty
                .GotoModule = String.Empty

                btnGo.PerformClick()

            End If
        End With


    End Sub

    Private Sub loadModules()
        Try
            'Dim dt As DataTable = MusterContainer.AppUser.ListModulesUserHasAccessTo(MusterContainer.AppUser.UserKey)
            Dim dt As DataTable = MusterContainer.AppUser.ListModulesUserCanSearch(MusterContainer.AppUser.UserKey)
            Dim dr As DataRow

            dr = dt.NewRow
            dr("PROPERTY_NAME") = "- Favorite Reports"
            dr("PROPERTY_ID") = -1
            dt.Rows.Add(dr)

            dr = dt.NewRow
            dr("PROPERTY_NAME") = "- All Reports"
            dr("PROPERTY_ID") = 0
            dt.Rows.Add(dr)

            dt.DefaultView.Sort = "PROPERTY_NAME"

            With cboModule
                .ValueMember = "PROPERTY_ID"
                .DisplayMember = "PROPERTY_NAME"
                .DataSource = dt.DefaultView
            End With


            'Dim oModProp As New MUSTER.BusinessLogic.pPropertyType("Modules")
            'Dim dtModules As DataTable
            'Dim drModule As DataRow
            'Dim i As Integer = 0

            'dtModules = oModProp.PropertiesTable

            'drModule = dtModules.NewRow
            'drModule.Item("Property Name") = "Favorite Reports"
            'drModule.Item("Property ID") = -1
            'drModule.Item("PropType_ID") = 0
            'dtModules.Rows.InsertAt(drModule, 0)

            'drModule = dtModules.NewRow
            'drModule.Item("Property Name") = "All Reports"
            'drModule.Item("Property ID") = 0
            'drModule.Item("PropType_ID") = 0
            'dtModules.Rows.InsertAt(drModule, dtModules.Rows.Count + 1)

            'dtModules.DefaultView.RowFilter = "PropType_ID=0 OR PropType_ID=" + oModProp.ID.ToString
            'Me.cboModule.DataSource = dtModules.DefaultView
            'Me.cboModule.DisplayMember = "Property Name"
            'Me.cboModule.ValueMember = "Property ID"

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub LoadFavReports()
        dtReports = oReports.GetReportsForUser(MusterContainer.AppUser.UserKey, , , True)
        If dtReports.Rows.Count <= 0 Then
            drReports = dtReports.NewRow
            drReports("REPORT_ID") = -1
            drReports("REPORT_NAME") = "No Favorite Reports"
            dtReports.Rows.Add(drReports)
        Else
            drReports = dtReports.NewRow
            drReports("REPORT_ID") = -1
            drReports("REPORT_NAME") = "- Please Select One -"
            dtReports.Rows.Add(drReports)
        End If
        dtReports.DefaultView.Sort = "REPORT_NAME"
        cboReports.DisplayMember = "REPORT_NAME"
        cboReports.ValueMember = "REPORT_ID"
        cboReports.DataSource = dtReports.DefaultView
        cboReports.SelectedValue = -1
    End Sub

    Private Sub MarkReportRowFav(ByVal reportID As Integer, ByVal isFav As Boolean)
        If Not dtReports Is Nothing Then
            If dtReports.Columns.Contains("FAV") Then
                For Each drReports In dtReports.Rows
                    If drReports("REPORT_ID") = reportID Then
                        drReports("FAV") = isFav
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub EnableDisableButtons(Optional ByVal bolbtnGo As Boolean = False, Optional ByVal bolbtnSaveAsFav As Boolean = False, Optional ByVal bolbtnDelAsFav As Boolean = False)
        If cboReports.DataSource Is Nothing Then
            btnGo.Enabled = False
            btnSaveasFavorite.Enabled = False
            btnDeleteFav.Enabled = False
        Else
            btnGo.Enabled = bolbtnGo
            btnSaveasFavorite.Enabled = bolbtnSaveAsFav
            btnDeleteFav.Enabled = bolbtnDelAsFav
        End If
    End Sub

    Private Function isReportSelectionValid() As Boolean
        Dim showError As Boolean = False
        If cboReports.SelectedValue Is Nothing Then
            showError = True
        ElseIf cboReports.SelectedValue = -1 Then
            showError = True
        End If
        If showError Then
            MsgBox("Please Select a Report")
            Return False
        Else
            Return True
        End If
    End Function
#End Region

#Region "Form Events"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Dim RptMgr As New InfoRepository.GenericDataManager
        Try
            'catDS = New DataSet
            'RptMgr.PopulateDataSet("SELECT Report_Name, Report_Loc FROM tblSYS_REPORT_MASTER WHERE ACTIVE = 1 order by report_id")

            'catDS = RptMgr.UniDataSet

            'cboReports.DataSource = catDS.Tables(0)
            'cboReports.DisplayMember = "Report_Name"

            CrystalReportViewer1.Height = Me.Height - 100
            CrystalReportViewer1.Width = Me.Width - 50

            CrystalReportViewer1.Visible = True
            loadModules()
            bolLoading = False

            LoadReportFromTickler()



        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ReportDisplay_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try
            CrystalReportViewer1.Height = Me.Height - 100
            CrystalReportViewer1.Width = Me.Width - 50

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "UI Events"
    Private Sub cboModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboModule.SelectedIndexChanged
        If bolLoading Then Exit Sub
        'Dim bolLoadingLocal As Boolean = bolLoading
        Try
            'bolLoading = True
            If cboModule.SelectedValue Is Nothing Then
                dtReports = New DataTable
                cboReports.DataSource = Nothing
                EnableDisableButtons()
            ElseIf cboModule.SelectedValue = -1 Then ' Fav Reports
                LoadFavReports()
            ElseIf cboModule.SelectedValue = 0 Then ' All Reports
                dtReports = oReports.GetReportsForUser(MusterContainer.AppUser.UserKey, , , )
                If dtReports.Rows.Count <= 0 Then
                    drReports = dtReports.NewRow
                    drReports("REPORT_ID") = -1
                    drReports("REPORT_NAME") = " - No Reports - "
                    If dtReports.Columns.Contains("FAV") Then
                        drReports("FAV") = False
                    End If
                    dtReports.Rows.Add(drReports)
                Else
                    drReports = dtReports.NewRow
                    drReports("REPORT_ID") = -1
                    drReports("REPORT_NAME") = "- Please Select One - "
                    If dtReports.Columns.Contains("FAV") Then
                        drReports("FAV") = False
                    End If
                    dtReports.Rows.Add(drReports)
                End If
                dtReports.DefaultView.Sort = "REPORT_NAME"
                cboReports.DisplayMember = "REPORT_NAME"
                cboReports.ValueMember = "REPORT_ID"
                cboReports.DataSource = dtReports.DefaultView
                'EnableDisableButtons()
                'bolLoading = False
                cboReports.SelectedValue = -1
            Else
                dtReports = oReports.GetReportsForUser(MusterContainer.AppUser.UserKey, cboModule.SelectedValue, , )
                If dtReports.Rows.Count <= 0 Then
                    drReports = dtReports.NewRow
                    drReports("REPORT_ID") = -1
                    drReports("REPORT_NAME") = " - No Reports - "
                    If dtReports.Columns.Contains("FAV") Then
                        drReports("FAV") = False
                    End If
                    dtReports.Rows.Add(drReports)
                Else
                    drReports = dtReports.NewRow
                    drReports("REPORT_ID") = -1
                    drReports("REPORT_NAME") = "- Please Select One - "
                    If dtReports.Columns.Contains("FAV") Then
                        drReports("FAV") = False
                    End If
                    dtReports.Rows.Add(drReports)
                End If
                dtReports.DefaultView.Sort = "REPORT_NAME"
                cboReports.DisplayMember = "REPORT_NAME"
                cboReports.ValueMember = "REPORT_ID"
                cboReports.DataSource = dtReports.DefaultView
                'EnableDisableButtons()
                'bolLoading = False
                cboReports.SelectedValue = -1
            End If

            'If cboModule.SelectedValue <> String.Empty Then '> 0 Or cboModule.SelectedValue = -1 Then
            '    mstContainer = Me.MdiParent
            '    Me.oReports = New MUSTER.BusinessLogic.pReport
            '    'oReports.Retrieve(mstContainer.AppUser.UserID)

            '    If cboModule.Items(cboModule.SelectedIndex).item(1) <> "Favorite Reports" Then
            '        dtReports = oReports.ListReportNames(Me.cboModule.SelectedValue, False, MusterContainer.AppUser.ID)
            '    Else
            '        dtReports = GetFavorites()
            '    End If
            '    'Me.cboModule.SelectedValue
            '    Me.btnDeleteFav.Visible = False
            '    Me.btnGo.Enabled = False
            '    Me.btnSaveasFavorite.Enabled = True
            '    Me.btnSaveasFavorite.Visible = False
            '    If dtReports.Rows.Count > 0 Then
            '        Me.btnGo.Enabled = True
            '        Me.btnSaveasFavorite.Visible = True
            '    Else
            '        Me.btnGo.Enabled = False
            '        Me.btnSaveasFavorite.Visible = False
            '    End If
            '    If cboModule.SelectedValue = -1 Then

            '        If dtReports.Rows.Count = 0 Then
            '            drReport = dtReports.NewRow
            '            drReport.Item("Report_Name") = "No Favorites"
            '            drReport.Item("Report_ID") = 0
            '            dtReports.Rows.Add(drReport)
            '        Else
            '            Me.btnDeleteFav.Visible = True
            '            Me.btnSaveasFavorite.Enabled = False
            '        End If
            '    End If
            '    Me.cboReports.DataSource = dtReports
            '    cboReports.DisplayMember = "REPORT_NAME"
            '    cboReports.ValueMember = "Report_ID"

            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
            'Finally
            'bolLoading = bolLoadingLocal
        End Try
    End Sub

    Private Sub cboReports_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboReports.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            If cboReports.DataSource Is Nothing Then
                EnableDisableButtons()
            ElseIf cboReports.SelectedValue Is Nothing Then
                EnableDisableButtons()
            ElseIf cboReports.SelectedValue <= 0 Then
                EnableDisableButtons()
            Else
                If dtReports Is Nothing Then
                    EnableDisableButtons()
                Else
                    If dtReports.Columns.Contains("FAV") Then
                        For Each drReports In dtReports.Rows
                            If drReports("REPORT_ID") = cboReports.SelectedValue Then
                                If drReports("FAV") Then
                                    EnableDisableButtons(True, False, True)
                                    Exit For
                                Else
                                    EnableDisableButtons(True, True, False)
                                    Exit For
                                End If
                            End If
                        Next
                    Else
                        EnableDisableButtons(True, True, True)
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnSaveasFavorite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveasFavorite.Click
        If Not isReportSelectionValid() Then Exit Sub

        'Dim oProfileInfo As MUSTER.Info.ProfileInfo
        'Dim strReportName As String
        'Dim dr As DataRowView
        Try

            oReports.SaveFavReport(MusterContainer.AppUser.UserKey, cboReports.SelectedValue, False, UIUtilsGen.ModuleID.[Global], MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            MarkReportRowFav(cboReports.SelectedValue, True)

            'If cboReports.Items.Count > 0 Then
            '    dr = cboReports.Items.Item(cboReports.SelectedIndex)
            '    strReportName = dr.Item("REPORT_NAME")
            '    If Not strReportName.StartsWith("No Favorites") Then
            '        'TODO  IsAFav property - Ask Adam what this is about JVC2 2/8/05
            '        oProfileInfo = oProfile.Retrieve(MusterContainer.AppUser.ID & "|FAVORITE REPORTS|" & strReportName & "|NONE")
            '        If oProfileInfo Is Nothing Then
            '            oProfileInfo = New MUSTER.Info.ProfileInfo(MusterContainer.AppUser.ID, "FAVORITE REPORTS", strReportName, "NONE", "NONE", False, MusterContainer.AppUser.ID, Today, MusterContainer.AppUser.ID, Today)
            '            oProfile.Add(oProfileInfo)
            '            oProfile.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
            '            If Not UIUtilsGen.HasRights(returnVal) Then
            '                Exit Sub
            '            End If
            '        Else
            '            oProfileInfo.Deleted = False
            '            oProfile.Add(oProfileInfo)
            '            oProfile.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
            '            If Not UIUtilsGen.HasRights(returnVal) Then
            '                Exit Sub
            '            End If
            '        End If
            '    End If
            'Else
            '    MsgBox("Please Select a Report")
            '    Exit Sub
            'End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnDeleteFav_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteFav.Click
        If Not isReportSelectionValid() Then Exit Sub

        'Dim oProfileInfo As MUSTER.Info.ProfileInfo
        'Dim oRpt As MUSTER.BusinessLogic.pReport
        'Dim strReportName As String
        'Dim dr As DataRowView
        Try
            oReports.SaveFavReport(MusterContainer.AppUser.UserKey, cboReports.SelectedValue, True, UIUtilsGen.ModuleID.[Global], MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            If Not cboModule.SelectedValue Is Nothing Then
                If cboModule.SelectedValue = -1 Then ' Fav Reports
                    LoadFavReports()
                Else
                    MarkReportRowFav(cboReports.SelectedValue, False)
                End If
            End If
            'If cboReports.Items.Count > 0 And Not cboReports.Text = String.Empty Then
            '    dr = cboReports.Items.Item(cboReports.SelectedIndex)
            '    strReportName = dr.Item("REPORT_NAME")

            '    'oRpt.Retrieve(strReportName)
            '    'oRpt.Deleted = False
            '    'oRpt.Save()
            '    'oReports.Update(oRpt)
            '    'oReports.SaveFavorites()

            '    oProfileInfo = oProfile.Retrieve(MusterContainer.AppUser.ID & "|FAVORITE REPORTS|" & strReportName & "|NONE")
            '    If Not oProfileInfo Is Nothing Then
            '        'oProfile.Remove(oProfileInfo)
            '        oProfileInfo.Deleted = True
            '        If oProfileInfo.User = String.Empty Then
            '            oProfileInfo.CreatedBy = MusterContainer.AppUser.ID
            '        Else
            '            oProfileInfo.ModifiedBy = MusterContainer.AppUser.ID
            '        End If
            '        oProfile.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
            '        If Not UIUtilsGen.HasRights(returnVal) Then
            '            Exit Sub
            '        End If

            '    End If

            '    'cboReports.Items.Remove(cboReports.Items.Item(cboReports.SelectedIndex))

            '    Me.cboModule_SelectedIndexChanged(sender, e)
            '    CrystalReportViewer1.ReportSource = Nothing
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Sub GenerateReport(Optional ByVal reportName As String = "", Optional ByVal params As Object() = Nothing, Optional ByVal promptForDoc As String = "", Optional ByVal PDFName As String = "", Optional ByVal moduleID As UIUtilsGen.ModuleID = UIUtilsGen.ModuleID.CAPProcess)

        'Validation 
        If reportName = String.Empty Then
            If Not isReportSelectionValid() Then Exit Sub
        End If
        If REPORT_PATH = "\" Then
            MsgBox("Invalid Report Location")
            Exit Sub
        End If

        ' Declarations
        Dim oParentFrm As MusterContainer = CType(Me.MdiParent, MusterContainer)
        Dim CrConnInfo As ConnectionInfo
        Dim CrTableLogon As TableLogOnInfo
        Dim crTable As CrystalDecisions.CrystalReports.Engine.Table
        Dim crTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim reportID As Integer = 0
        Dim crDatabase As CrystalDecisions.CrystalReports.Engine.Database
        Dim rname As String
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor


        'Get Report Path
        If reportName.Length > 0 Then
            Panel1.Visible = False
            Panel1.Height = 0
            CrystalReportViewer1.Location = New Point(0, 0)
        End If

        If reportName = String.Empty AndAlso cboReports.Items.Count > 0 Then
            reportID = cboReports.SelectedValue
        End If

        oReports.Retrieve(IIf(reportName = String.Empty, reportID, reportName))
        rname = oReports.Path
        oRpt = New ReportDocument()
        'oRpt.

        If Not rname.ToUpper.StartsWith(REPORT_PATH.ToUpper) Then
            If Not rname.StartsWith("\") Then
                rname = REPORT_PATH + rname
            End If
        End If

        'Load Report     
        'Assign DB Information To Report & Sub Reports
        Try
            If System.IO.File.Exists(rname) Then
                oRpt.Load(rname) ', OpenReportMethod.OpenReportByTempCopy
                oRpt.DataSourceConnections.Clear()


                ' DB Connection Details

                ' To  Get Db Connection Details from App Config
                'RptCon.DatabaseName = ConfigurationManager.AppSettings("DatabaseName").ToString()
                'RptCon.ServerName = ConfigurationManager.AppSettings("ServerName").ToString()
                'RptCon.UserID = ConfigurationManager.AppSettings("Username").ToString()
                'RptCon.Password = ConfigurationManager.AppSettings("Password").ToString()

                ' Get DB Connection Details From Current Login details
                Dim LocalUserSettings As Microsoft.Win32.Registry
                Dim lastLogin As Object = LocalUserSettings.CurrentUser.GetValue("MusterSQLConnection")
                Dim LoginServer, LoginDatabase, LoginUser, LoginPwd As String

                For Each str As String In lastLogin.ToString.Split(";")
                    If str.StartsWith("Data Source") Then
                        LoginServer = str.Split("=")(1)
                    ElseIf str.StartsWith("Initial Catalog") Then
                        LoginDatabase = str.Split("=")(1)
                    ElseIf str.StartsWith("user") Then
                        LoginUser = str.Split("=")(1)
                    ElseIf str.StartsWith("password") Then
                        LoginPwd = str.Split("=")(1)
                    End If
                Next
                ' Assigning connection details to Report Login
                Dim RptCon As New ConnectionInfo()

                RptCon.ServerName = LoginServer.Replace("'", "").Trim()
                RptCon.DatabaseName = LoginDatabase.Replace("'", "").Trim()
                RptCon.UserID = LoginUser.Replace("'", "").Trim()
                RptCon.Password = LoginPwd.Replace("'", "").Trim()
                For Each table As CrystalDecisions.CrystalReports.Engine.Table In oRpt.Database.Tables
                    AssignTableConnection(table, RptCon)
                Next
                ' Now loop through all the sections and its objects to do the same for the subreports
                '
                For Each section As CrystalDecisions.CrystalReports.Engine.Section In oRpt.ReportDefinition.Sections
                    ' In each section we need to loop through all the reporting objects
                    For Each reportObject As CrystalDecisions.CrystalReports.Engine.ReportObject In section.ReportObjects

                        If reportObject.Kind = ReportObjectKind.SubreportObject Then

                            Dim subReport As SubreportObject = DirectCast(reportObject, SubreportObject)
                            Dim subDocument As ReportDocument = subReport.OpenSubreport(subReport.SubreportName)

                            For Each table As CrystalDecisions.CrystalReports.Engine.Table In subDocument.Database.Tables
                                AssignTableConnection(table, RptCon)
                            Next

                            subDocument.SetDatabaseLogon(RptCon.UserID, RptCon.Password, RptCon.ServerName, RptCon.DatabaseName)

                        End If
                    Next
                Next
                oRpt.SetDatabaseLogon(RptCon.UserID, RptCon.Password, RptCon.ServerName, RptCon.DatabaseName)
                CrystalReportViewer1.Visible = True
                'Report Parameter
                Try
                    'If there are report parameters then load the parameter form
                    Dim nSubReportParamCount As Integer = 0
                    If (oRpt.DataDefinition.ParameterFields.Count = 0) Or _
                            (oRpt.DataDefinition.ParameterFields.Count <= nSubReportParamCount) Then
                        CrystalReportViewer1.ReportSource = oRpt

                    ElseIf reportName.Length > 0 Then
                        oParameterForm = New CrystalReportsParamForm(params)
                        oParameterForm.Cr = oRpt
                        oParameterForm.Report = oReports.Retrieve(IIf(reportName = String.Empty, reportID, reportName))
                        CrystalReportViewer1.Tag = "HOLD"
                        oParameterForm.PushParamters()
                        oParameterForm.Dispose()
                    Else
                        oParameterForm = New CrystalReportsParamForm
                        oParameterForm.Cr = oRpt
                        oParameterForm.Report = oReports.Retrieve(IIf(reportName = String.Empty, cboReports.SelectedValue, reportName))
                        oParameterForm.ShowDialog()
                        If FlagParamDialogClosed Then
                            FlagParamDialogClosed = False
                            ''oRpt.Dispose()
                            Me.Refresh()
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                        oParameterForm.Dispose()

                    End If

                    Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

                    If promptForDoc.Length > 0 AndAlso promptForDoc <> "PROCESSED" AndAlso MsgBox(promptForDoc, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then


                        Dim nLen As Integer
                        Dim strPhysicalPath As String
                        Dim fInfo As New System.IO.FileInfo(rname)
                        nLen = (fInfo.Name.Length - 4)
                        Dim strPdfFileName As String = fInfo.Name.Remove(nLen, 4)
                        Dim strDirName As String = fInfo.DirectoryName
                        Dim CrystalExportOptions As ExportOptions
                        Dim CrystalDiskFileDestinationOptions As DiskFileDestinationOptions
                        Dim procYear As String = String.Empty
                        Dim procOwner As String = String.Empty
                        Dim naming As String = String.Empty

                        If PDFName.Length = 0 Then
                            If Not params Is Nothing AndAlso params.GetUpperBound(0) >= 3 Then
                                procYear = "_" & params(0).ToString

                                If (params(3).ToString.Length > 0) Then procOwner = "_" & params(3)

                            End If

                            naming = String.Format("{0}{1}", procYear, procOwner)
                        Else
                            naming = String.Format("_{0}", PDFName)
                        End If

                        CrystalDiskFileDestinationOptions = New DiskFileDestinationOptions

                        strPhysicalPath = Me.TEMPLETTER_PATH & "\" & strPdfFileName & naming & ".PDF"
                        CrystalDiskFileDestinationOptions.DiskFileName = strPhysicalPath
                        CrystalExportOptions = oRpt.ExportOptions
                        With CrystalExportOptions
                            .DestinationOptions = CrystalDiskFileDestinationOptions
                            .ExportDestinationType = ExportDestinationType.DiskFile
                            .ExportFormatType = ExportFormatType.PortableDocFormat
                        End With
                        oRpt.Export()

                        UIUtilsGen.SaveDocument(0, 0, strPdfFileName & naming & ".PDF", reportName, DOC_PATH, reportName & " for " & naming, moduleID, 0, 0, 0)

                        UIUtilsGen.OpenInPDFFile(strPhysicalPath)
                    ElseIf promptForDoc <> "PROCESSED" Then
                        If Not Me.Visible Then
                            Me.MdiParent = _container
                            Me.CrystalReportViewer1.ReportSource = oRpt
                            Me.Show()
                        End If
                    Else
                        MsgBox("Report Completed. Please open the report viewer to see it/")
                    End If
                Catch ex As Exception
                    Dim MyErr As New ErrorReport(ex)
                    MyErr.ShowDialog()
                    'MsgBox("Error : " + ex.Message + vbCrLf + "In : " + ex.Source.ToString, MsgBoxStyle.OKOnly, "Error Producing Report")
                End Try

                Me.Cursor = System.Windows.Forms.Cursors.Default
                CrystalReportViewer1.ReportSource = oRpt
                CrystalReportViewer1.Refresh()
                oRpt.DataSourceConnections.Clear()

            Else
                If Not oRpt Is Nothing Then
                    MsgBox("the report " & cboReports.Text & " is not found.")
                    Me.Cursor = System.Windows.Forms.Cursors.Default
                    Exit Sub
                End If
            End If

        Catch engEx As LogOnException
            MessageBox.Show _
           ("Incorrect Logon Parameters. Check your user name and password.")
        Catch engEx As DataSourceException
            MessageBox.Show _
            ("An error has occurred while connecting to the database.")
        Catch engEx As EngineException
            MessageBox.Show(engEx.Message)

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error : " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub AssignTableConnection(ByVal table As CrystalDecisions.CrystalReports.Engine.Table, ByVal connection As ConnectionInfo)
        ' Cache the logon info block
        Dim logOnInfo As TableLogOnInfo = table.LogOnInfo

        connection.Type = logOnInfo.ConnectionInfo.Type

        ' Set the connection
        logOnInfo.ConnectionInfo = connection

        ' Apply the connection to the table!

        table.LogOnInfo.ConnectionInfo.DatabaseName = connection.DatabaseName
        table.LogOnInfo.ConnectionInfo.ServerName = connection.ServerName
        table.LogOnInfo.ConnectionInfo.UserID = connection.UserID
        table.LogOnInfo.ConnectionInfo.Password = connection.Password
        table.LogOnInfo.ConnectionInfo.Type = connection.Type
        table.ApplyLogOnInfo(logOnInfo)
    End Sub
    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
        GenerateReport()
    End Sub
#End Region

#Region "External Events"
    Private Sub ReturnReport() Handles oParameterForm.ReturnReport
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        Try

            If CrystalReportViewer1.Tag <> "HOLD" Then
                CrystalReportViewer1.ReportSource = oRpt
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try

    End Sub
#End Region


#Region "Miscellaneous code"

    Private Sub LoadReport()

        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table
        Dim TableCounter

        'If you are using a Strongly Typed report (Imported in 
        'your project) named CrystalReport1.rpt use the 
        'following: 

        'If you are using a Non-Typed report, and 
        'loading a report outside of the project, use the 
        'following: 

        Dim crReportDocument As New ReportDocument
        crReportDocument.Load("c:\myReports\myReport.rpt")

        'Set the ConnectionInfo properties for logging on to 
        'the Database 

        'If you are using ODBC, this should be the 
        'DSN name NOT the physical server name. If 
        'you are NOT using ODBC, this should be the 
        'physical server name 

        With crConnectionInfo
            .ServerName = "DSN or Server Name"

            'If you are connecting to Oracle there is no 
            'DatabaseName. Use an empty string. 
            'For example, .DatabaseName = "" 

            .DatabaseName = "DatabaseName"
            .UserID = "Your User ID"
            .Password = "Your Password"
        End With

        'This code works for both user tables and stored 
        'procedures. Set the CrTables to the Tables collection 
        'of the report 

        CrTables = crReportDocument.Database.Tables

        'Loop through each table in the report and apply the 
        'LogonInfo information 

        For Each CrTable In CrTables
            crtableLogoninfo = CrTable.LogOnInfo
            crtableLogoninfo.ConnectionInfo = crConnectionInfo
            CrTable.ApplyLogOnInfo(crtableLogoninfo)

            'If your DatabaseName is changing at runtime, specify 
            'the table location. 
            'For example, when you are reporting off of a 
            'Northwind database on SQL server you 
            'should have the following line of code: 

            CrTable.Location = "Northwind.dbo." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
        Next

        'Set the viewer to the report object to be previewed. 
        CrystalReportViewer1.ReportSource = crReportDocument
    End Sub
#End Region

    Private Sub CrystalReportViewer1_ReportRefresh(ByVal source As Object, ByVal e As CrystalDecisions.Windows.Forms.ViewerEventArgs) Handles CrystalReportViewer1.ReportRefresh

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load

    End Sub

    Private Sub lblReport_Click(sender As Object, e As EventArgs) Handles lblReport.Click

    End Sub

    Private Sub lblModule_Click(sender As Object, e As EventArgs) Handles lblModule.Click

    End Sub
End Class
