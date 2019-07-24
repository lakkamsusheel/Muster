Public Class RescheduleInspection
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private bolLoading As Boolean = False
    Private dtRescheduledDate As Date
    Private strRescheduledTime As String = String.Empty
    Private WithEvents oInspection As MUSTER.BusinessLogic.pInspection
    Private dtMyFacilities As DataTable
    Friend CallingForm As Form
    Dim returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(ByRef oInspec As MUSTER.BusinessLogic.pInspection, ByVal dtMyFacs As DataTable)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        Cursor.Current = Cursors.AppStarting
        bolLoading = True
        oInspection = oInspec
        dtMyFacilities = dtMyFacs
        'If Date.Compare(oInspection.RescheduledDate, CDate("01/01/0001")) <> 0 Then
        dtRescheduledDate = oInspection.RescheduledDate
        'End If
        strRescheduledTime = oInspection.RescheduledTime
        LoadSchedule()
        bolLoading = False
        Cursor.Current = Cursors.Default
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
    Friend WithEvents pnlRescheduleBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlRescheduleDetails As System.Windows.Forms.Panel
    Friend WithEvents lblScheduleDate As System.Windows.Forms.Label
    Friend WithEvents lblScheduleTime As System.Windows.Forms.Label
    Friend WithEvents dtPickerScheduleDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents cmbTime As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlRescheduleBottom = New System.Windows.Forms.Panel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlRescheduleDetails = New System.Windows.Forms.Panel
        Me.dtPickerScheduleDate = New System.Windows.Forms.DateTimePicker
        Me.lblScheduleTime = New System.Windows.Forms.Label
        Me.lblScheduleDate = New System.Windows.Forms.Label
        Me.cmbTime = New System.Windows.Forms.ComboBox
        Me.pnlRescheduleBottom.SuspendLayout()
        Me.pnlRescheduleDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlRescheduleBottom
        '
        Me.pnlRescheduleBottom.Controls.Add(Me.btnDelete)
        Me.pnlRescheduleBottom.Controls.Add(Me.btnCancel)
        Me.pnlRescheduleBottom.Controls.Add(Me.btnSave)
        Me.pnlRescheduleBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlRescheduleBottom.Location = New System.Drawing.Point(0, 86)
        Me.pnlRescheduleBottom.Name = "pnlRescheduleBottom"
        Me.pnlRescheduleBottom.Size = New System.Drawing.Size(360, 40)
        Me.pnlRescheduleBottom.TabIndex = 0
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(232, 8)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(152, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(72, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        '
        'pnlRescheduleDetails
        '
        Me.pnlRescheduleDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlRescheduleDetails.Controls.Add(Me.cmbTime)
        Me.pnlRescheduleDetails.Controls.Add(Me.dtPickerScheduleDate)
        Me.pnlRescheduleDetails.Controls.Add(Me.lblScheduleTime)
        Me.pnlRescheduleDetails.Controls.Add(Me.lblScheduleDate)
        Me.pnlRescheduleDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlRescheduleDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlRescheduleDetails.Name = "pnlRescheduleDetails"
        Me.pnlRescheduleDetails.Size = New System.Drawing.Size(360, 86)
        Me.pnlRescheduleDetails.TabIndex = 1
        '
        'dtPickerScheduleDate
        '
        Me.dtPickerScheduleDate.Checked = False
        Me.dtPickerScheduleDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickerScheduleDate.Location = New System.Drawing.Point(128, 8)
        Me.dtPickerScheduleDate.Name = "dtPickerScheduleDate"
        Me.dtPickerScheduleDate.ShowCheckBox = True
        Me.dtPickerScheduleDate.Size = New System.Drawing.Size(104, 20)
        Me.dtPickerScheduleDate.TabIndex = 2
        '
        'lblScheduleTime
        '
        Me.lblScheduleTime.Location = New System.Drawing.Point(42, 32)
        Me.lblScheduleTime.Name = "lblScheduleTime"
        Me.lblScheduleTime.Size = New System.Drawing.Size(83, 17)
        Me.lblScheduleTime.TabIndex = 1
        Me.lblScheduleTime.Text = "Schedule Time:"
        '
        'lblScheduleDate
        '
        Me.lblScheduleDate.Location = New System.Drawing.Point(44, 8)
        Me.lblScheduleDate.Name = "lblScheduleDate"
        Me.lblScheduleDate.Size = New System.Drawing.Size(83, 17)
        Me.lblScheduleDate.TabIndex = 0
        Me.lblScheduleDate.Text = "Schedule Date:"
        '
        'cmbTime
        '
        Me.cmbTime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTime.Location = New System.Drawing.Point(128, 32)
        Me.cmbTime.Name = "cmbTime"
        Me.cmbTime.Size = New System.Drawing.Size(104, 21)
        Me.cmbTime.TabIndex = 13
        '
        'RescheduleInspection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(360, 126)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlRescheduleDetails)
        Me.Controls.Add(Me.pnlRescheduleBottom)
        Me.Name = "RescheduleInspection"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Schedule Inspection"
        Me.pnlRescheduleBottom.ResumeLayout(False)
        Me.pnlRescheduleDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Support Routines"
    Private Sub LoadSchedule()
        cmbTime.DataSource = oInspection.GetInspectionTimes.Tables(0)
        cmbTime.DisplayMember = "PROPERTY_NAME"
        cmbTime.ValueMember = "PROPERTY_NAME"

        If oInspection.ID > 0 Then
            btnDelete.Enabled = True
            If Date.Compare(oInspection.RescheduledDate, CDate("01/01/0001")) = 0 Then
                UIUtilsGen.SetDatePickerValue(dtPickerScheduleDate, oInspection.ScheduledDate.ToShortDateString)
            Else
                UIUtilsGen.SetDatePickerValue(dtPickerScheduleDate, oInspection.RescheduledDate.ToShortDateString)
            End If
            If oInspection.RescheduledTime = String.Empty Then
                UIUtilsGen.SetComboboxItemByText(cmbTime, oInspection.ScheduledTime)
            Else
                UIUtilsGen.SetComboboxItemByText(cmbTime, oInspection.RescheduledTime)
            End If
            ' #2897 if the current owner is different for the inspection, user can only delete inspection
            For Each dr As DataRow In dtMyFacilities.Select("INSPECTION_ID = " + oInspection.ID.ToString)
                If dr("OWNER_ID") = oInspection.OwnerID Then
                    Exit For
                Else
                    ' user can only delete
                    Me.Text = "Delete Inspection"
                    dtPickerScheduleDate.Enabled = False
                    cmbTime.Enabled = False
                    btnSave.Enabled = False
                End If
            Next
        Else
            btnDelete.Enabled = False
            UIUtilsGen.CreateEmptyFormatDatePicker(dtPickerScheduleDate)
            cmbTime.SelectedIndex = 0
        End If
        btnSave.Enabled = oInspection.IsDirty
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub dtPickerScheduleDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickerScheduleDate.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickerScheduleDate)
            If oInspection.ID <= 0 Then
                'UIUtilsGen.FillDateobjectValues(oInspection.ScheduledDate, dtPickerScheduleDate.Text)
                oInspection.ScheduledDate = UIUtilsGen.GetDatePickerValue(dtPickerScheduleDate).Date
            Else
                If Date.Compare(dtRescheduledDate, CDate("01/01/0001")) = 0 Then
                    'UIUtilsGen.FillDateobjectValues(oInspection.RescheduledDate, dtPickerScheduleDate.Text)
                    oInspection.RescheduledDate = UIUtilsGen.GetDatePickerValue(dtPickerScheduleDate).Date
                Else
                    'UIUtilsGen.FillDateobjectValues(oInspection.ScheduledDate, dtRescheduledDate)
                    'UIUtilsGen.FillDateobjectValues(oInspection.RescheduledDate, dtPickerScheduleDate.Text)
                    oInspection.ScheduledDate = dtRescheduledDate
                    oInspection.RescheduledDate = UIUtilsGen.GetDatePickerValue(dtPickerScheduleDate).Date
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbTime_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTime.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            If oInspection.ID <= 0 Then
                oInspection.ScheduledTime = UIUtilsGen.GetComboBoxText(cmbTime)
            Else
                If strRescheduledTime = String.Empty Then
                    oInspection.RescheduledTime = UIUtilsGen.GetComboBoxText(cmbTime)
                Else
                    oInspection.ScheduledTime = strRescheduledTime
                    oInspection.RescheduledTime = UIUtilsGen.GetComboBoxText(cmbTime)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim drRow As DataRow
        Dim strTime, strDate, strErr As String
        Try
            strErr = String.Empty
            strTime = String.Empty
            strDate = String.Empty
            If Date.Compare(UIUtilsGen.GetDatePickerValue(dtPickerScheduleDate), CDate("01/01/0001")) = 0 Then
                strErr = lblScheduleDate.Text.Trim.TrimEnd(":")
            End If
            If cmbTime.SelectedValue Is Nothing Then
                strTime = String.Empty
            Else
                strTime = UIUtilsGen.GetComboBoxText(cmbTime)
            End If
            If strTime = String.Empty Then
                If strErr <> String.Empty Then strErr += ", "
                strErr += lblScheduleTime.Text.Trim.TrimEnd(":")
            End If
            If strErr <> String.Empty Then
                MsgBox("Required: " + strErr)
                Exit Sub
            End If
            Cursor.Current = Cursors.AppStarting
            For Each drRow In dtMyFacilities.Rows
                If Integer.Parse(drRow("OWNER_ID")) <> oInspection.OwnerID Then
                    If Not drRow.Item("RESCHEDULED DATE") Is System.DBNull.Value Then
                        If Date.Compare(CDate(drRow.Item("RESCHEDULED DATE")), dtPickerScheduleDate.Value.ToShortDateString) = 0 _
                            And ((IIf(drRow.Item("RESCHEDULED TIME") Is DBNull.Value, String.Empty, drRow.Item("RESCHEDULED TIME")) = UIUtilsGen.GetComboBoxText(cmbTime)) _
                            And (drRow.Item("SCHEDULED BY") = MusterContainer.AppUser.Name)) Then
                            Cursor.Current = Cursors.Default
                            MsgBox("Facility " + drRow("FACILITY ID").ToString + " for owner " + drRow("OWNER NAME") + " is Scheduled on " + dtPickerScheduleDate.Value.ToShortDateString + " . Please select an other Scheduled date and time")
                            'oInspection.Reset()
                            Exit Sub
                        End If
                    ElseIf Not drRow.Item("SCHEDULED DATE") Is System.DBNull.Value And drRow.Item("RESCHEDULED DATE") Is System.DBNull.Value Then
                        If Date.Compare(CDate(drRow.Item("SCHEDULED DATE")), dtPickerScheduleDate.Value.ToShortDateString) = 0 _
                           And ((IIf(drRow.Item("SCHEDULED TIME") Is DBNull.Value, String.Empty, drRow.Item("SCHEDULED TIME")) = UIUtilsGen.GetComboBoxText(cmbTime)) _
                           And (drRow.Item("SCHEDULED BY") = MusterContainer.AppUser.Name)) Then
                            Cursor.Current = Cursors.Default
                            MsgBox("Facility " + drRow("FACILITY ID").ToString + " for owner " + drRow("OWNER NAME") + " is Scheduled on " + dtPickerScheduleDate.Value.ToShortDateString + " . Please select an other Scheduled date and time")
                            'oInspection.Reset()
                            Exit Sub
                        End If
                    End If
                End If
            Next
            Cursor.Current = Cursors.Default
            oInspection.ScheduledBy = MusterContainer.AppUser.ID
            'oInspection.StaffID = MusterContainer.AppUser.UserKey
            oInspection.InspectionType = 988 ' COMPLIANCE AUDIT
            oInspection.SubmittedDate = CDate("01/01/0001")

            If oInspection.ID <= 0 Then
                oInspection.CreatedBy = MusterContainer.AppUser.ID
            Else
                oInspection.ModifiedBy = MusterContainer.AppUser.ID
            End If

            Dim Success As Boolean = False
            Success = oInspection.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If Success Then
                ' commented per bug 2363
                'MsgBox("Inspection Schedule saved")
                CallingForm.Tag = "1"
                Me.Close()
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim strErr As String = String.Empty
        Try
            If Date.Compare(oInspection.CheckListGenDate, CDate("01/01/0001")) <> 0 Then
                strErr = "There is Checklist associated with the inspection" + vbCrLf
            End If
            Dim Results As Long = MsgBox(strErr + "Are you sure you want to delete the inspection?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Delete Inspection Confirmation")
            If Results = MsgBoxResult.Yes Then
                ' according to ddd - clear scheduled date, rescheduled date, time, letter generated and checklist generated
                ' since each inspection in unique, cannot reuse the id. hence, delete the row (instance)
                If oInspection.ID < 0 Then
                    oInspection.Reset()
                    CallingForm.Tag = "0"
                Else
                    oInspection.Deleted = True
                    If oInspection.ID <= 0 Then
                        oInspection.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oInspection.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oInspection.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, True, True)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    CallingForm.Tag = "2"
                End If
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            oInspection.Reset()
            CallingForm.Tag = "0"
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            If oInspection.IsDirty Then
                Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                If Results = MsgBoxResult.Yes Then
                    btnSave.PerformClick()
                ElseIf Results = MsgBoxResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                Else
                    oInspection.Reset()
                End If
            End If
            oInspection.Remove(oInspection.ID)
            oInspection = New MUSTER.BusinessLogic.pInspection
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "External Events"
    Private Sub InspectionChanged(ByVal bolValue As Boolean) Handles oInspection.evtInspectionChanged
        btnSave.Enabled = bolValue Or oInspection.IsDirty
    End Sub
    'Private Sub InspectionColChanged(ByVal bolValue As Boolean) Handles oInspection.evtColChanged
    '    btnSave.Enabled = bolValue Or oInspection.IsDirty
    'End Sub
#End Region
End Class
