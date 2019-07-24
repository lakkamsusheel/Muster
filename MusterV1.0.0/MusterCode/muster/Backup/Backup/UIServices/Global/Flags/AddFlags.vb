Public Class AddFlags
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Friend AddModifyFlag As Integer = 0
    Dim flagID, entityID, entityType, calID As Integer
    Dim strModule, strEntitySequenceNum As String
    Dim pCalLocal As MUSTER.BusinessLogic.pCalendar
    Dim Selectedrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim pFlagLocal As MUSTER.BusinessLogic.pFlag
    Dim returnVal As String = String.Empty
#End Region
    Friend Event FlagAdded() ' Used to tell the parent that a flag was added.
    Friend Event RefreshCalendar()
#Region " Windows Form Designer generated code "

    Private mcontainer As MusterContainer

    Public Sub New(Optional ByVal entity_ID As Integer = 0, Optional ByVal entity_Type As Integer = 0, Optional ByVal [Module] As String = "", Optional ByVal flag_ID As Integer = 0, Optional ByVal calendarID As Integer = 0, Optional ByVal Selrow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing, Optional ByVal entitySequenceNum As String = "", Optional ByRef container As MusterContainer = Nothing)


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
        strEntitySequenceNum = entitySequenceNum
        flagID = flag_ID
        calID = calendarID
        Selectedrow = Selrow
        If flagID = 0 Then
            Me.Text = "Add Flag for "
        Else
            Me.Text = "Modify Flag for "
        End If
        lblEntityName.Text = Me.Text
        Select Case entityType
            Case UIUtilsGen.EntityTypes.Owner
                Me.Text += "Owner (" + entityID.ToString + ") in " + strModule
                lblEntityName.Text += "Owner (" + entityID.ToString + ")"
            Case UIUtilsGen.EntityTypes.Facility
                Me.Text += "Facility (" + entityID.ToString + ") in " + strModule
                lblEntityName.Text += "Facility (" + entityID.ToString + ")"
            Case UIUtilsGen.EntityTypes.Company
                Me.Text += "Company (" + entityID.ToString + ") in " + strModule
                lblEntityName.Text += "Company (" + entityID.ToString + ")"
            Case UIUtilsGen.EntityTypes.Licensee
                Me.Text += "Licensee (" + entityID.ToString + ") in " + strModule
                lblEntityName.Text += "Licensee (" + entityID.ToString + ")"
            Case UIUtilsGen.EntityTypes.LUST_Event
                Me.Text += "Lust Event (" + IIf(strEntitySequenceNum = "", entityID.ToString, strEntitySequenceNum) + ") in " + strModule
                lblEntityName.Text += "Lust Event (" + IIf(strEntitySequenceNum = "", entityID.ToString, strEntitySequenceNum) + ")"
        End Select
        pFlagLocal = New MUSTER.BusinessLogic.pFlag
        pCalLocal = New MUSTER.BusinessLogic.pCalendar
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
    Friend WithEvents lblEntityName As System.Windows.Forms.Label
    Friend WithEvents lblFlag As System.Windows.Forms.Label
    Friend WithEvents chkCCEntry As System.Windows.Forms.CheckBox
    Friend WithEvents txtFlag As System.Windows.Forms.TextBox
    Friend WithEvents btnSaveFlag As System.Windows.Forms.Button
    Friend WithEvents rbFlagDueToMe As System.Windows.Forms.RadioButton
    Friend WithEvents rbFlagToDo As System.Windows.Forms.RadioButton
    Friend WithEvents grpFlagType As System.Windows.Forms.GroupBox
    Friend WithEvents lblDueDate As System.Windows.Forms.Label
    Friend WithEvents dtPickFlagDueDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblWhenToTurnRed As System.Windows.Forms.Label
    Friend WithEvents dtPickWhenToTurnRed As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblFlag = New System.Windows.Forms.Label
        Me.lblEntityName = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSaveFlag = New System.Windows.Forms.Button
        Me.txtFlag = New System.Windows.Forms.TextBox
        Me.chkCCEntry = New System.Windows.Forms.CheckBox
        Me.grpFlagType = New System.Windows.Forms.GroupBox
        Me.lblDueDate = New System.Windows.Forms.Label
        Me.dtPickFlagDueDate = New System.Windows.Forms.DateTimePicker
        Me.rbFlagToDo = New System.Windows.Forms.RadioButton
        Me.rbFlagDueToMe = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblWhenToTurnRed = New System.Windows.Forms.Label
        Me.dtPickWhenToTurnRed = New System.Windows.Forms.DateTimePicker
        Me.grpFlagType.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFlag
        '
        Me.lblFlag.Location = New System.Drawing.Point(16, 40)
        Me.lblFlag.Name = "lblFlag"
        Me.lblFlag.Size = New System.Drawing.Size(88, 16)
        Me.lblFlag.TabIndex = 7
        Me.lblFlag.Text = "Flag Description"
        '
        'lblEntityName
        '
        Me.lblEntityName.Location = New System.Drawing.Point(16, 16)
        Me.lblEntityName.Name = "lblEntityName"
        Me.lblEntityName.Size = New System.Drawing.Size(320, 16)
        Me.lblEntityName.TabIndex = 6
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(192, 358)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        '
        'btnSaveFlag
        '
        Me.btnSaveFlag.Location = New System.Drawing.Point(80, 358)
        Me.btnSaveFlag.Name = "btnSaveFlag"
        Me.btnSaveFlag.Size = New System.Drawing.Size(96, 23)
        Me.btnSaveFlag.TabIndex = 3
        Me.btnSaveFlag.Text = "Save Flag"
        '
        'txtFlag
        '
        Me.txtFlag.Location = New System.Drawing.Point(17, 57)
        Me.txtFlag.Multiline = True
        Me.txtFlag.Name = "txtFlag"
        Me.txtFlag.Size = New System.Drawing.Size(315, 167)
        Me.txtFlag.TabIndex = 0
        Me.txtFlag.Text = ""
        '
        'chkCCEntry
        '
        Me.chkCCEntry.Location = New System.Drawing.Point(18, 230)
        Me.chkCCEntry.Name = "chkCCEntry"
        Me.chkCCEntry.Size = New System.Drawing.Size(144, 24)
        Me.chkCCEntry.TabIndex = 1
        Me.chkCCEntry.Text = "Create Calendar Entry"
        '
        'grpFlagType
        '
        Me.grpFlagType.Controls.Add(Me.lblDueDate)
        Me.grpFlagType.Controls.Add(Me.dtPickFlagDueDate)
        Me.grpFlagType.Controls.Add(Me.rbFlagToDo)
        Me.grpFlagType.Controls.Add(Me.rbFlagDueToMe)
        Me.grpFlagType.Controls.Add(Me.Label1)
        Me.grpFlagType.Location = New System.Drawing.Point(79, 256)
        Me.grpFlagType.Name = "grpFlagType"
        Me.grpFlagType.Size = New System.Drawing.Size(195, 90)
        Me.grpFlagType.TabIndex = 2
        Me.grpFlagType.TabStop = False
        Me.grpFlagType.Visible = False
        '
        'lblDueDate
        '
        Me.lblDueDate.Location = New System.Drawing.Point(8, 14)
        Me.lblDueDate.Name = "lblDueDate"
        Me.lblDueDate.Size = New System.Drawing.Size(56, 24)
        Me.lblDueDate.TabIndex = 3
        Me.lblDueDate.Text = "Due Date"
        Me.lblDueDate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'dtPickFlagDueDate
        '
        Me.dtPickFlagDueDate.Checked = False
        Me.dtPickFlagDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickFlagDueDate.Location = New System.Drawing.Point(72, 16)
        Me.dtPickFlagDueDate.Name = "dtPickFlagDueDate"
        Me.dtPickFlagDueDate.ShowCheckBox = True
        Me.dtPickFlagDueDate.Size = New System.Drawing.Size(96, 20)
        Me.dtPickFlagDueDate.TabIndex = 0
        '
        'rbFlagToDo
        '
        Me.rbFlagToDo.Location = New System.Drawing.Point(72, 68)
        Me.rbFlagToDo.Name = "rbFlagToDo"
        Me.rbFlagToDo.Size = New System.Drawing.Size(120, 16)
        Me.rbFlagToDo.TabIndex = 2
        Me.rbFlagToDo.Text = "To Do"
        '
        'rbFlagDueToMe
        '
        Me.rbFlagDueToMe.Checked = True
        Me.rbFlagDueToMe.Location = New System.Drawing.Point(72, 48)
        Me.rbFlagDueToMe.Name = "rbFlagDueToMe"
        Me.rbFlagDueToMe.Size = New System.Drawing.Size(120, 16)
        Me.rbFlagDueToMe.TabIndex = 1
        Me.rbFlagDueToMe.TabStop = True
        Me.rbFlagDueToMe.Text = "Due To Me"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 32)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Calendar Type"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblWhenToTurnRed
        '
        Me.lblWhenToTurnRed.Location = New System.Drawing.Point(162, 228)
        Me.lblWhenToTurnRed.Name = "lblWhenToTurnRed"
        Me.lblWhenToTurnRed.Size = New System.Drawing.Size(71, 24)
        Me.lblWhenToTurnRed.TabIndex = 9
        Me.lblWhenToTurnRed.Text = "Flag Turns Red on"
        Me.lblWhenToTurnRed.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtPickWhenToTurnRed
        '
        Me.dtPickWhenToTurnRed.Checked = False
        Me.dtPickWhenToTurnRed.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickWhenToTurnRed.Location = New System.Drawing.Point(237, 230)
        Me.dtPickWhenToTurnRed.Name = "dtPickWhenToTurnRed"
        Me.dtPickWhenToTurnRed.ShowCheckBox = True
        Me.dtPickWhenToTurnRed.Size = New System.Drawing.Size(97, 20)
        Me.dtPickWhenToTurnRed.TabIndex = 8
        '
        'AddFlags
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(352, 393)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblWhenToTurnRed)
        Me.Controls.Add(Me.dtPickWhenToTurnRed)
        Me.Controls.Add(Me.txtFlag)
        Me.Controls.Add(Me.grpFlagType)
        Me.Controls.Add(Me.chkCCEntry)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSaveFlag)
        Me.Controls.Add(Me.lblEntityName)
        Me.Controls.Add(Me.lblFlag)
        Me.Name = "AddFlags"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Flag"
        Me.grpFlagType.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub AddFlags_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If flagID > 0 Then
                ' modify flag
                If Selectedrow Is Nothing Then
                    MsgBox("Error loading flag")
                    Me.Close()
                    Exit Sub
                End If
                txtFlag.Text = Selectedrow.Cells("DESCRIPTION").Value
                If Selectedrow.Cells("TURNS RED ON").Value Is DBNull.Value Then
                    UIUtilsGen.SetDatePickerValue(dtPickWhenToTurnRed, CDate("01/01/0001"))
                Else
                    UIUtilsGen.SetDatePickerValue(dtPickWhenToTurnRed, Selectedrow.Cells("TURNS RED ON").Value)
                End If
                If calID > 0 Then
                    chkCCEntry.Checked = Not chkCCEntry.Checked
                    chkCCEntry.Checked = True
                    pCalLocal.Retrieve(calID)
                    UIUtilsGen.SetDatePickerValue(dtPickFlagDueDate, pCalLocal.DateDue)
                    If pCalLocal.ToDo Then
                        rbFlagToDo.Checked = True
                    Else
                        rbFlagDueToMe.Checked = True
                    End If
                Else
                    chkCCEntry.Checked = Not chkCCEntry.Checked
                    chkCCEntry.Checked = False
                End If
            Else
                chkCCEntry.Checked = Not chkCCEntry.Checked
                chkCCEntry.Checked = False
                UIUtilsGen.SetDatePickerValue(dtPickWhenToTurnRed, DateAdd(DateInterval.Day, 120, Now.Date))
            End If
            txtFlag.Focus()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSaveFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveFlag.Click
        Dim strErr As String
        Try
            If txtFlag.Text.Trim.Length <= 0 Then
                strErr = "Flag Description, "
            End If
            If chkCCEntry.Checked Then
                If Date.Compare(UIUtilsGen.GetDatePickerValue(dtPickFlagDueDate), CDate("01/01/0001")) = 0 Then
                    strErr += "Due Date, "
                End If
                If rbFlagDueToMe.Checked = False And rbFlagToDo.Checked = False Then
                    strErr += "Calendar Type"
                End If
            End If
            If strErr <> String.Empty Then
                MsgBox("Please enter the following: " + strErr.Trim.TrimEnd(","))
                Exit Sub
            End If
            If chkCCEntry.Checked Then
                If calID > 0 Then
                    pCalLocal.Retrieve(calID)
                    pCalLocal.NotificationDate = UIUtilsGen.GetDatePickerValue(dtPickFlagDueDate)
                    pCalLocal.DateDue = UIUtilsGen.GetDatePickerValue(dtPickFlagDueDate)
                    pCalLocal.TaskDescription = txtFlag.Text
                    pCalLocal.DueToMe = rbFlagDueToMe.Checked
                    pCalLocal.ToDo = rbFlagToDo.Checked
                Else
                    Dim oCalInfo As New MUSTER.Info.CalendarInfo(0, _
                                 UIUtilsGen.GetDatePickerValue(dtPickFlagDueDate), _
                                 UIUtilsGen.GetDatePickerValue(dtPickFlagDueDate), _
                                 0, _
                                 Me.txtFlag.Text, _
                                 MusterContainer.AppUser.ID, _
                                 MusterContainer.AppUser.ID, _
                                 String.Empty, _
                                 Me.rbFlagDueToMe.Checked, _
                                 Me.rbFlagToDo.Checked, _
                                 False, _
                                 False, _
                                 MusterContainer.AppUser.ID, _
                                 CDate("01/01/0001"), _
                                 String.Empty, _
                                 CDate("01/01/0001"), _
                                 entityType, _
                                 entityID)
                    'oCalInfo.OwningEntityID = entityID
                    'oCalInfo.OwningEntityType = entityType
                    pCalLocal.Add(oCalInfo)
                End If
                pCalLocal.Save()
                calID = pCalLocal.CalendarId
                RaiseEvent RefreshCalendar()
            Else
                If calID > 0 Then
                    pCalLocal.Retrieve(calID)
                    pCalLocal.Deleted = True
                    pCalLocal.ModifiedBy = MusterContainer.AppUser.ID
                    pCalLocal.Save()
                    calID = 0
                    RaiseEvent RefreshCalendar()
                End If
            End If
            If flagID > 0 Then
                pFlagLocal.Retrieve(flagID)
                pFlagLocal.FlagDescription = txtFlag.Text
                pFlagLocal.DueDate = UIUtilsGen.GetDatePickerValue(dtPickFlagDueDate)
                pFlagLocal.TurnsRedOn = UIUtilsGen.GetDatePickerValue(dtPickWhenToTurnRed)
                pFlagLocal.CalendarInfoID = calID
                pFlagLocal.ModifiedBy = MusterContainer.AppUser.ID
                'pFlagLocal.SourceUserID = MusterContainer.AppUser.ID
            Else
                Dim oFlagInfo As New MUSTER.Info.FlagInfo(0, _
                                entityID, _
                                entityType, _
                                txtFlag.Text, _
                                False, _
                                UIUtilsGen.GetDatePickerValue(dtPickFlagDueDate), _
                                strModule, _
                                calID, _
                                MusterContainer.AppUser.ID, _
                                CDate("01/01/0001"), _
                                String.Empty, _
                                CDate("01/01/0001"), _
                                UIUtilsGen.GetDatePickerValue(dtPickWhenToTurnRed), _
                                MusterContainer.AppUser.ID)
                pFlagLocal.Add(oFlagInfo)
            End If
            pFlagLocal.Save(CType(UIUtilsGen.ModuleID.Global, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            RaiseEvent FlagAdded()
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        If Not (pCalLocal Is Nothing) Then
            pCalLocal.Reset()
        End If
        If Not (pFlagLocal Is Nothing) Then
            pFlagLocal.Reset()
        End If
        Me.Close()
    End Sub
    Private Sub chkCCEntry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCCEntry.CheckedChanged
        If chkCCEntry.Checked Then
            grpFlagType.Visible = True
            Me.Height = 420
            btnCancel.Top = 358
            btnSaveFlag.Top = 358
        Else
            grpFlagType.Visible = False
            Me.Height = 328
            btnCancel.Top = 264
            btnSaveFlag.Top = 264
        End If
        Me.Refresh()
    End Sub

    Private Sub ShowFlags_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        If Not mcontainer Is Nothing Then
            mcontainer.HoldClosing = True
        End If
    End Sub


    Private Sub dtPickWhenToTurnRed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickWhenToTurnRed.ValueChanged
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickWhenToTurnRed)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickFlagDueDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickFlagDueDate.ValueChanged
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickWhenToTurnRed)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
