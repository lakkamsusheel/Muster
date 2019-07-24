'-------------------------------------------------------------------------------
' MUSTER.MUSTER.ShowFlags
'   Provides the mechanism for makingFlag entries.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        ??      8/??/04    Original class definition.
'
'  1.1        MR     1/31/05    Added Functions to Save Calendar and Flag
'-------------------------------------------------------------------------------
Imports System
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic



Public Class ShowFlags
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Dim result As DialogResult
    Dim WithEvents ACfrm As AddFlags
    Private WithEvents pFlagLocal As MUSTER.BusinessLogic.pFlag
    Dim bolLoading As Boolean
    Public entityID, entityType, eventID, eventType As Integer
    Public strModule, strParentFormText, strEntitySequenceNum As String
    Dim returnVal As String = String.Empty
#End Region
#Region "User Defined Events"
    ' Used to notify parent that a flag was added
    Friend Event FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String)
    Friend Event RefreshCalendar()
#End Region

    Private mcontainer As MusterContainer
#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByVal entity_ID As Integer = 0, Optional ByVal entity_Type As Integer = 0, Optional ByVal [Module] As String = "", Optional ByVal overrideShowAllModules As Boolean = False, Optional ByVal enableShowAllModules As Boolean = True, Optional ByVal entitySequenceNum As String = "", Optional ByVal event_ID As Integer = 0, Optional ByVal event_Type As Integer = 0, Optional ByRef container As MusterContainer = Nothing)
        MyBase.New()

        mcontainer = container

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        '
        ' Getting the Current Entity Details.
        '
        strModule = [Module]
        entityID = entity_ID
        entityType = entity_Type
        eventID = event_ID
        eventType = event_Type
        strEntitySequenceNum = entitySequenceNum
        ProcessFormText(overrideShowAllModules, strEntitySequenceNum)
        chkShowAllModules.Enabled = enableShowAllModules
        pFlagLocal = New MUSTER.BusinessLogic.pFlag
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
    Friend WithEvents lblEntityName As System.Windows.Forms.Label
    Friend WithEvents lblFlags As System.Windows.Forms.Label
    Friend WithEvents pnlFlagsTop As System.Windows.Forms.Panel
    Friend WithEvents pnlFlagsGrid As System.Windows.Forms.Panel
    Friend WithEvents pnlFlagsBottom As System.Windows.Forms.Panel
    Friend WithEvents ugFlag As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnModifyFlag As System.Windows.Forms.Button
    Friend WithEvents btnFlagAdd As System.Windows.Forms.Button
    Friend WithEvents btnFlagCancel As System.Windows.Forms.Button
    Friend WithEvents btnFlagDelete As System.Windows.Forms.Button
    Friend WithEvents chkShowAllModules As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlFlagsTop = New System.Windows.Forms.Panel
        Me.lblEntityName = New System.Windows.Forms.Label
        Me.lblFlags = New System.Windows.Forms.Label
        Me.pnlFlagsGrid = New System.Windows.Forms.Panel
        Me.ugFlag = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFlagsBottom = New System.Windows.Forms.Panel
        Me.btnModifyFlag = New System.Windows.Forms.Button
        Me.btnFlagAdd = New System.Windows.Forms.Button
        Me.btnFlagCancel = New System.Windows.Forms.Button
        Me.btnFlagDelete = New System.Windows.Forms.Button
        Me.chkShowAllModules = New System.Windows.Forms.CheckBox
        Me.pnlFlagsTop.SuspendLayout()
        Me.pnlFlagsGrid.SuspendLayout()
        CType(Me.ugFlag, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFlagsBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlFlagsTop
        '
        Me.pnlFlagsTop.Controls.Add(Me.lblEntityName)
        Me.pnlFlagsTop.Controls.Add(Me.lblFlags)
        Me.pnlFlagsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFlagsTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlFlagsTop.Name = "pnlFlagsTop"
        Me.pnlFlagsTop.Size = New System.Drawing.Size(720, 40)
        Me.pnlFlagsTop.TabIndex = 0
        '
        'lblEntityName
        '
        Me.lblEntityName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEntityName.Location = New System.Drawing.Point(80, 8)
        Me.lblEntityName.Name = "lblEntityName"
        Me.lblEntityName.Size = New System.Drawing.Size(552, 16)
        Me.lblEntityName.TabIndex = 114
        '
        'lblFlags
        '
        Me.lblFlags.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFlags.Location = New System.Drawing.Point(8, 8)
        Me.lblFlags.Name = "lblFlags"
        Me.lblFlags.Size = New System.Drawing.Size(64, 16)
        Me.lblFlags.TabIndex = 113
        Me.lblFlags.Text = "Flags for:"
        '
        'pnlFlagsGrid
        '
        Me.pnlFlagsGrid.Controls.Add(Me.ugFlag)
        Me.pnlFlagsGrid.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFlagsGrid.Location = New System.Drawing.Point(0, 40)
        Me.pnlFlagsGrid.Name = "pnlFlagsGrid"
        Me.pnlFlagsGrid.Size = New System.Drawing.Size(720, 311)
        Me.pnlFlagsGrid.TabIndex = 2
        '
        'ugFlag
        '
        Me.ugFlag.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFlag.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFlag.Location = New System.Drawing.Point(0, 0)
        Me.ugFlag.Name = "ugFlag"
        Me.ugFlag.Size = New System.Drawing.Size(720, 311)
        Me.ugFlag.TabIndex = 122
        Me.ugFlag.Text = "Flags"
        '
        'pnlFlagsBottom
        '
        Me.pnlFlagsBottom.Controls.Add(Me.btnModifyFlag)
        Me.pnlFlagsBottom.Controls.Add(Me.btnFlagAdd)
        Me.pnlFlagsBottom.Controls.Add(Me.btnFlagCancel)
        Me.pnlFlagsBottom.Controls.Add(Me.btnFlagDelete)
        Me.pnlFlagsBottom.Controls.Add(Me.chkShowAllModules)
        Me.pnlFlagsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFlagsBottom.Location = New System.Drawing.Point(0, 351)
        Me.pnlFlagsBottom.Name = "pnlFlagsBottom"
        Me.pnlFlagsBottom.Size = New System.Drawing.Size(720, 42)
        Me.pnlFlagsBottom.TabIndex = 121
        '
        'btnModifyFlag
        '
        Me.btnModifyFlag.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnModifyFlag.Location = New System.Drawing.Point(112, 8)
        Me.btnModifyFlag.Name = "btnModifyFlag"
        Me.btnModifyFlag.Size = New System.Drawing.Size(96, 26)
        Me.btnModifyFlag.TabIndex = 123
        Me.btnModifyFlag.Text = "Modify Flag"
        '
        'btnFlagAdd
        '
        Me.btnFlagAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFlagAdd.Location = New System.Drawing.Point(8, 8)
        Me.btnFlagAdd.Name = "btnFlagAdd"
        Me.btnFlagAdd.Size = New System.Drawing.Size(96, 26)
        Me.btnFlagAdd.TabIndex = 119
        Me.btnFlagAdd.Text = "Add Flag"
        '
        'btnFlagCancel
        '
        Me.btnFlagCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFlagCancel.Location = New System.Drawing.Point(320, 8)
        Me.btnFlagCancel.Name = "btnFlagCancel"
        Me.btnFlagCancel.Size = New System.Drawing.Size(75, 26)
        Me.btnFlagCancel.TabIndex = 120
        Me.btnFlagCancel.Text = "Close"
        '
        'btnFlagDelete
        '
        Me.btnFlagDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFlagDelete.Location = New System.Drawing.Point(216, 8)
        Me.btnFlagDelete.Name = "btnFlagDelete"
        Me.btnFlagDelete.Size = New System.Drawing.Size(96, 26)
        Me.btnFlagDelete.TabIndex = 121
        Me.btnFlagDelete.Text = "Delete Flag"
        '
        'chkShowAllModules
        '
        Me.chkShowAllModules.Location = New System.Drawing.Point(403, 11)
        Me.chkShowAllModules.Name = "chkShowAllModules"
        Me.chkShowAllModules.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkShowAllModules.Size = New System.Drawing.Size(120, 20)
        Me.chkShowAllModules.TabIndex = 122
        Me.chkShowAllModules.Text = "Show All Modules"
        Me.chkShowAllModules.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ShowFlags
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(720, 393)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlFlagsBottom)
        Me.Controls.Add(Me.pnlFlagsGrid)
        Me.Controls.Add(Me.pnlFlagsTop)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Location = New System.Drawing.Point(100, 200)
        Me.Name = "ShowFlags"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Show Flags"
        Me.pnlFlagsTop.ResumeLayout(False)
        Me.pnlFlagsGrid.ResumeLayout(False)
        CType(Me.ugFlag, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFlagsBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub ShowFlags_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ugFlag.DataSource = pFlagLocal.GetFlagsDS(entityID, entityType, False, IIf(chkShowAllModules.Checked, String.Empty, strModule))
            SetupGrid()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFlagAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFlagAdd.Click
        Try
            If IsNothing(ACfrm) Then
                ACfrm = New AddFlags(entityID, entityType, strModule, , , , strEntitySequenceNum, mcontainer)
                AddHandler ACfrm.Closing, AddressOf frmAddFlagClosing
                AddHandler ACfrm.Closed, AddressOf frmAddFlagClosed
            End If
            ACfrm.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub frmAddFlagClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            ugFlag.DataSource = pFlagLocal.GetFlagsDS(entityID, entityType, False, IIf(chkShowAllModules.Checked, String.Empty, strModule))
            SetupGrid()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub frmAddFlagClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        ACfrm = Nothing
    End Sub
    Private Sub btnFlagCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFlagCancel.Click
        Me.Close()
    End Sub
    Private Sub btnFlagDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFlagDelete.Click
        Try
            If ugFlag.Rows.Count <= 0 Then Exit Sub

            If ugFlag.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete")
                Exit Sub
            Else
                result = MessageBox.Show("Are you Sure you want to Delete this Record?", "Flags", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then
                    Exit Sub
                Else
                    If Not (ugFlag.ActiveRow.Cells("CALENDAR_INFO_ID").Value Is DBNull.Value) Then
                        If ugFlag.ActiveRow.Cells("CALENDAR_INFO_ID").Value <> 0 Then
                            'Remove the Calendar Entry Bfore Removing Flag.
                            MusterContainer.pCalendar.Retrieve(Integer.Parse(ugFlag.ActiveRow.Cells("CALENDAR_INFO_ID").Text))
                            If MusterContainer.pCalendar.CalendarId = ugFlag.ActiveRow.Cells("CALENDAR_INFO_ID").Value Then
                                MusterContainer.pCalendar.Deleted = True
                                MusterContainer.pCalendar.Save()
                                MusterContainer.pCalendar.Remove(Integer.Parse(ugFlag.ActiveRow.Cells("CALENDAR_INFO_ID").Text))
                                ' refresh calendar
                                RaiseEvent RefreshCalendar()
                            End If
                        End If
                    End If
                    MusterContainer.pFlag.Retrieve(Integer.Parse(ugFlag.ActiveRow.Cells("FLAG_ID").Value))
                    MusterContainer.pFlag.Deleted = True
                    If MusterContainer.pFlag.ID <= 0 Then
                        MusterContainer.pFlag.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        MusterContainer.pFlag.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    MusterContainer.pFlag.Save(CType(UIUtilsGen.ModuleID.Global, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    MusterContainer.pFlag.Remove(Integer.Parse(ugFlag.ActiveRow.Cells("FLAG_ID").Value))
                    ugFlag.ActiveRow.Delete(False)
                    RaiseEvent FlagAdded(entityID, entityType, strModule, strParentFormText)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkShowAllModules_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAllModules.CheckedChanged
        Try
            If bolLoading Then Exit Sub
            If chkShowAllModules.Checked = True Then
                ugFlag.DataSource = MusterContainer.pFlag.GetFlagsDS(entityID, entityType, False)
                ProcessFormText(True)
            Else
                ugFlag.DataSource = MusterContainer.pFlag.GetFlagsDS(entityID, entityType, False, strModule)
                ProcessFormText(True)
            End If
            SetupGrid()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnModifyFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifyFlag.Click
        Try
            If Not ugFlag.Rows.Count > 0 Then Exit Sub
            If ugFlag.ActiveRow Is Nothing Then
                MsgBox("Select row to Modify.")
                Exit Sub
            End If
            If IsNothing(ACfrm) Then
                Dim calID As Integer
                If Not (ugFlag.ActiveRow.Cells("CALENDAR_INFO_ID").Value Is DBNull.Value) Then
                    calID = ugFlag.ActiveRow.Cells("CALENDAR_INFO_ID").Value
                End If
                ACfrm = New AddFlags(entityID, entityType, ugFlag.ActiveRow.Cells("MODULE").Value, ugFlag.ActiveRow.Cells("FLAG_ID").Value, calID, ugFlag.ActiveRow, strEntitySequenceNum, mcontainer)
                AddHandler ACfrm.Closing, AddressOf frmAddFlagClosing
                AddHandler ACfrm.Closed, AddressOf frmAddFlagClosed
            End If
            ACfrm.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFlag_AfterSelectChange(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles ugFlag.AfterSelectChange
        Try
            ProcessRowTest()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ShowFlags_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try
            pnlFlagsGrid.Height = Me.Height - pnlFlagsTop.Height - pnlFlagsBottom.Height - 27
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFlag_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFlag.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ProcessRowTest()
            If btnModifyFlag.Enabled Then
                btnModifyFlag_Click(sender, e)
            Else
                MsgBox("You Cannot Modify the selected row")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ACfrm_FlagAdded() Handles ACfrm.FlagAdded
        Try
            ugFlag.DataSource = MusterContainer.pFlag.GetFlagsDS(entityID, entityType, False, IIf(chkShowAllModules.Checked, String.Empty, strModule))
            SetupGrid()
            RaiseEvent FlagAdded(entityID, entityType, strModule, strParentFormText)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ACfrm_RefreshCalendar() Handles ACfrm.RefreshCalendar
        RaiseEvent RefreshCalendar()
    End Sub
    Private Sub ProcessRowTest()
        Dim dtSuperUsers As DataTable
        Dim drow As DataRow
        Dim bolEnableFlag As Boolean = False
        Try
            With ugFlag.Selected
                If .Rows.Count > 0 Then
                    If MusterContainer.AppUser.ID = ugFlag.ActiveRow.Cells("USER ID").Value Then
                        btnFlagDelete.Enabled = True
                        btnModifyFlag.Enabled = True
                        Exit Sub
                    Else
                        dtSuperUsers = MusterContainer.AppUser.ListSupervisedUsers()
                        If dtSuperUsers.Rows.Count > 0 Then
                            For Each drow In dtSuperUsers.Rows
                                If drow("USER_ID") = ugFlag.ActiveRow.Cells("USER ID").Value Then
                                    bolEnableFlag = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Else
                    If ugFlag.Rows.Count > 0 Then
                        If MusterContainer.AppUser.ID = ugFlag.Rows(0).Cells("USER ID").Value Then
                            btnFlagDelete.Enabled = True
                            btnModifyFlag.Enabled = True
                            Exit Sub
                        Else
                            dtSuperUsers = MusterContainer.AppUser.ListSupervisedUsers()
                            If dtSuperUsers.Rows.Count > 0 Then
                                For Each drow In dtSuperUsers.Rows
                                    If drow("USER_ID") = ugFlag.Rows(0).Cells("USER ID").Value Then
                                        bolEnableFlag = True
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End With
            If bolEnableFlag Then
                btnFlagDelete.Enabled = True
                btnModifyFlag.Enabled = True
            Else
                btnFlagDelete.Enabled = False
                btnModifyFlag.Enabled = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SetupGrid()
        Try
            ugFlag.DisplayLayout.Bands(0).Columns("FLAG_ID").Hidden = True
            ugFlag.DisplayLayout.Bands(0).Columns("FLAG_COLOR").Hidden = True
            ugFlag.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
            ugFlag.DisplayLayout.Bands(0).Columns("CREATED_BY").Hidden = True
            ugFlag.DisplayLayout.Bands(0).Columns("LAST_EDITED_BY").Hidden = True
            ugFlag.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True
            ugFlag.DisplayLayout.Bands(0).Columns("CALENDAR_INFO_ID").Hidden = True
            ugFlag.DisplayLayout.Bands(0).Columns("ENTITY_TYPE").Hidden = True

            ugFlag.DisplayLayout.Bands(0).Columns("CREATED ON").Width = 90
            ugFlag.DisplayLayout.Bands(0).Columns("DESCRIPTION").Width = 275
            ugFlag.DisplayLayout.Bands(0).Columns("DESCRIPTION").CellMultiLine = Infragistics.Win.DefaultableBoolean.True
            ugFlag.DisplayLayout.Bands(0).Columns("DESCRIPTION").VertScrollBar = True
            ugFlag.DisplayLayout.Bands(0).Columns("DESCRIPTION").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True
            ugFlag.DisplayLayout.Bands(0).Columns("MODULE").Width = 80
            ugFlag.DisplayLayout.Bands(0).Columns("USER ID").Width = 75
            ugFlag.DisplayLayout.Bands(0).Columns("DUE DATE").Width = 90

            ugFlag.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
            ugFlag.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
            ugFlag.DisplayLayout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
            ugFlag.DisplayLayout.Bands(0).Override.RowSizingAutoMaxLines = 5

            ugFlag.DisplayLayout.Bands(0).Columns("CREATED ON").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugFlag.DisplayLayout.Bands(0).Columns("DUE DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugFlag.DisplayLayout.Bands(0).Columns("TURNS RED ON").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugFlag.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugFlag.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            ugFlag.DisplayLayout.Bands(0).Columns("ENTITY TYPE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugFlag.DisplayLayout.Bands(0).Columns("MODULE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugFlag.DisplayLayout.Bands(0).Columns("USER ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugFlag.DisplayLayout.Bands(0).Columns("DUE DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugFlag.DisplayLayout.Bands(0).Columns("ENTITY ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            ugFlag.DisplayLayout.Bands(0).Columns("TURNS RED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            If ugFlag.Rows.Count > 0 Then
                ProcessRowTest()
                ' change flag color here
                Dim rowColor As System.Drawing.Color
                For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugFlag.Rows
                    Select Case ugRow.Cells("FLAG_COLOR").Value
                        Case "YELLOW"
                            rowColor = Color.Yellow
                        Case "RED"
                            rowColor = Color.Red
                        Case Else
                            rowColor = Color.White
                    End Select
                    ugRow.Appearance.BackColor = rowColor
                Next
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ProcessFormText(Optional ByVal overrideShowAllModules As Boolean = False, Optional ByVal entitySequenceNum As String = "")
        Try
            Select Case entityType
                Case UIUtilsGen.EntityTypes.Owner
                    lblEntityName.Text = "Owner : " + entityID.ToString
                    Me.Text = "Flags for Owner (" + entityID.ToString + ") in " + IIf(chkShowAllModules.Checked, "All Modules", strModule)
                Case UIUtilsGen.EntityTypes.Facility
                    lblEntityName.Text = "Facility : " + entityID.ToString
                    If Not overrideShowAllModules Then
                        bolLoading = True
                        chkShowAllModules.Checked = True
                        bolLoading = False
                    End If
                    Me.Text = "Flags for Facility (" + entityID.ToString + ") in " + IIf(chkShowAllModules.Checked, "All Modules", strModule)
                Case UIUtilsGen.EntityTypes.Company
                    lblEntityName.Text = "Company : " + entityID.ToString
                    Me.Text = "Flags for Company (" + entityID.ToString + ") in " + IIf(chkShowAllModules.Checked, "All Modules", strModule)
                Case UIUtilsGen.EntityTypes.Licensee
                    lblEntityName.Text = "Licensee : " + entityID.ToString
                    Me.Text = "Flags for Licensee (" + entityID.ToString + ") in " + IIf(chkShowAllModules.Checked, "All Modules", strModule)
                Case UIUtilsGen.EntityTypes.Inspection
                    lblEntityName.Text = "Inspection : " + entityID.ToString
                    Me.Text = "Flags for Inspection (" + entityID.ToString + ") in " + IIf(chkShowAllModules.Checked, "All Modules", strModule)
                Case UIUtilsGen.EntityTypes.LUST_Event
                    Me.Text = "Flags for Lust Event (" + IIf(entitySequenceNum = "", entityID.ToString, entitySequenceNum) + ") in " + IIf(chkShowAllModules.Checked, "All Modules", strModule)
                Case UIUtilsGen.EntityTypes.FinancialEvent
                    lblEntityName.Text = "Financial Event : " + IIf(entitySequenceNum = "", entityID.ToString, entitySequenceNum)
                    Me.Text = "Flags for Financial Event (" + IIf(entitySequenceNum = "", entityID.ToString, entitySequenceNum) + ") in " + IIf(chkShowAllModules.Checked, "All Modules", strModule)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ShowFlags_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        If Not mcontainer Is Nothing Then
            mcontainer.HoldClosing = True
        End If
    End Sub
End Class
