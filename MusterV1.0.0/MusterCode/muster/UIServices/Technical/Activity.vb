Public Class Activity
    Inherits System.Windows.Forms.Form

#Region " Local Variables "

    Friend CallingForm As Form
    Friend Mode As Int16
    Friend EventActivityID As Int64
    Friend TFStatus As Int16

    Private bolLoading As Boolean
    Private oLocalLustEvent As MUSTER.BusinessLogic.pLustEvent
    Private WithEvents oLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
    Private WithEvents oComments As New MUSTER.BusinessLogic.pComments
    Private dtComments As DataTable
    Private nTotalOpenDocs As Integer
    Dim returnVal As String = String.Empty
    Dim bolUncompletedDoc As Boolean = False

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal oLustEvent As MUSTER.BusinessLogic.pLustEvent)
        MyBase.New()

        oLocalLustEvent = oLustEvent

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
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cmbActivity As System.Windows.Forms.ComboBox
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents dtPickStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickCloseDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtPickCompletedDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblCloseDate As System.Windows.Forms.Label
    Friend WithEvents lblCompletedDate As System.Windows.Forms.Label
    Friend WithEvents gbGWSSamples As System.Windows.Forms.GroupBox
    Friend WithEvents dt2ndGWSBelow As System.Windows.Forms.DateTimePicker
    Friend WithEvents dt1stGWSBelow As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btAddSystem As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbActivity = New System.Windows.Forms.ComboBox
        Me.lblCloseDate = New System.Windows.Forms.Label
        Me.lblCompletedDate = New System.Windows.Forms.Label
        Me.lblActivity = New System.Windows.Forms.Label
        Me.lblStartDate = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.lblComments = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.dtPickStartDate = New System.Windows.Forms.DateTimePicker
        Me.dtPickCloseDate = New System.Windows.Forms.DateTimePicker
        Me.dtPickCompletedDate = New System.Windows.Forms.DateTimePicker
        Me.gbGWSSamples = New System.Windows.Forms.GroupBox
        Me.dt2ndGWSBelow = New System.Windows.Forms.DateTimePicker
        Me.dt1stGWSBelow = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btAddSystem = New System.Windows.Forms.Button
        Me.gbGWSSamples.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbActivity
        '
        Me.cmbActivity.Location = New System.Drawing.Point(160, 24)
        Me.cmbActivity.Name = "cmbActivity"
        Me.cmbActivity.Size = New System.Drawing.Size(248, 21)
        Me.cmbActivity.TabIndex = 0
        '
        'lblCloseDate
        '
        Me.lblCloseDate.Location = New System.Drawing.Point(64, 120)
        Me.lblCloseDate.Name = "lblCloseDate"
        Me.lblCloseDate.Size = New System.Drawing.Size(88, 16)
        Me.lblCloseDate.TabIndex = 158
        Me.lblCloseDate.Text = "Closed Date:"
        Me.lblCloseDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCompletedDate
        '
        Me.lblCompletedDate.Location = New System.Drawing.Point(8, 88)
        Me.lblCompletedDate.Name = "lblCompletedDate"
        Me.lblCompletedDate.Size = New System.Drawing.Size(152, 16)
        Me.lblCompletedDate.TabIndex = 157
        Me.lblCompletedDate.Text = "Technically Completed Date:"
        Me.lblCompletedDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblActivity
        '
        Me.lblActivity.Location = New System.Drawing.Point(80, 24)
        Me.lblActivity.Name = "lblActivity"
        Me.lblActivity.Size = New System.Drawing.Size(64, 16)
        Me.lblActivity.TabIndex = 156
        Me.lblActivity.Text = "Activity:"
        Me.lblActivity.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblStartDate
        '
        Me.lblStartDate.Location = New System.Drawing.Point(72, 56)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(80, 16)
        Me.lblStartDate.TabIndex = 155
        Me.lblStartDate.Text = "Start Date: "
        Me.lblStartDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(160, 152)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(344, 72)
        Me.txtComments.TabIndex = 4
        Me.txtComments.Text = ""
        '
        'lblComments
        '
        Me.lblComments.Location = New System.Drawing.Point(72, 160)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(80, 16)
        Me.lblComments.TabIndex = 161
        Me.lblComments.Text = "Comments:"
        Me.lblComments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(408, 248)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(96, 26)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(208, 248)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(96, 26)
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Save"
        '
        'dtPickStartDate
        '
        Me.dtPickStartDate.Checked = False
        Me.dtPickStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickStartDate.Location = New System.Drawing.Point(160, 56)
        Me.dtPickStartDate.Name = "dtPickStartDate"
        Me.dtPickStartDate.ShowCheckBox = True
        Me.dtPickStartDate.Size = New System.Drawing.Size(104, 20)
        Me.dtPickStartDate.TabIndex = 1
        '
        'dtPickCloseDate
        '
        Me.dtPickCloseDate.Checked = False
        Me.dtPickCloseDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickCloseDate.Location = New System.Drawing.Point(160, 120)
        Me.dtPickCloseDate.Name = "dtPickCloseDate"
        Me.dtPickCloseDate.ShowCheckBox = True
        Me.dtPickCloseDate.Size = New System.Drawing.Size(104, 20)
        Me.dtPickCloseDate.TabIndex = 3
        '
        'dtPickCompletedDate
        '
        Me.dtPickCompletedDate.Checked = False
        Me.dtPickCompletedDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickCompletedDate.Location = New System.Drawing.Point(160, 88)
        Me.dtPickCompletedDate.Name = "dtPickCompletedDate"
        Me.dtPickCompletedDate.ShowCheckBox = True
        Me.dtPickCompletedDate.Size = New System.Drawing.Size(104, 20)
        Me.dtPickCompletedDate.TabIndex = 2
        '
        'gbGWSSamples
        '
        Me.gbGWSSamples.Controls.Add(Me.dt2ndGWSBelow)
        Me.gbGWSSamples.Controls.Add(Me.dt1stGWSBelow)
        Me.gbGWSSamples.Controls.Add(Me.Label1)
        Me.gbGWSSamples.Controls.Add(Me.Label2)
        Me.gbGWSSamples.Location = New System.Drawing.Point(280, 48)
        Me.gbGWSSamples.Name = "gbGWSSamples"
        Me.gbGWSSamples.Size = New System.Drawing.Size(224, 80)
        Me.gbGWSSamples.TabIndex = 162
        Me.gbGWSSamples.TabStop = False
        '
        'dt2ndGWSBelow
        '
        Me.dt2ndGWSBelow.Checked = False
        Me.dt2ndGWSBelow.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dt2ndGWSBelow.Location = New System.Drawing.Point(112, 48)
        Me.dt2ndGWSBelow.Name = "dt2ndGWSBelow"
        Me.dt2ndGWSBelow.ShowCheckBox = True
        Me.dt2ndGWSBelow.Size = New System.Drawing.Size(104, 20)
        Me.dt2ndGWSBelow.TabIndex = 159
        '
        'dt1stGWSBelow
        '
        Me.dt1stGWSBelow.Checked = False
        Me.dt1stGWSBelow.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dt1stGWSBelow.Location = New System.Drawing.Point(112, 16)
        Me.dt1stGWSBelow.Name = "dt1stGWSBelow"
        Me.dt1stGWSBelow.ShowCheckBox = True
        Me.dt1stGWSBelow.Size = New System.Drawing.Size(104, 20)
        Me.dt1stGWSBelow.TabIndex = 158
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 161
        Me.Label1.Text = "2nd GWS Below"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 160
        Me.Label2.Text = "1st GWS Below"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btAddSystem
        '
        Me.btAddSystem.Location = New System.Drawing.Point(8, 248)
        Me.btAddSystem.Name = "btAddSystem"
        Me.btAddSystem.Size = New System.Drawing.Size(144, 26)
        Me.btAddSystem.TabIndex = 163
        Me.btAddSystem.Text = "Add Remediation System"
        Me.btAddSystem.Visible = False
        '
        'Activity
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(552, 294)
        Me.Controls.Add(Me.btAddSystem)
        Me.Controls.Add(Me.gbGWSSamples)
        Me.Controls.Add(Me.dtPickCompletedDate)
        Me.Controls.Add(Me.dtPickCloseDate)
        Me.Controls.Add(Me.dtPickStartDate)
        Me.Controls.Add(Me.txtComments)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.lblComments)
        Me.Controls.Add(Me.cmbActivity)
        Me.Controls.Add(Me.lblCloseDate)
        Me.Controls.Add(Me.lblCompletedDate)
        Me.Controls.Add(Me.lblActivity)
        Me.Controls.Add(Me.lblStartDate)
        Me.Name = "Activity"
        Me.Text = "Add / Modify Activity"
        Me.gbGWSSamples.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Event Operations "
    Private Sub Activity_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProcessLoad()

    End Sub

    Private Sub ProcessLoad()
        Dim ds As DataSet
        Dim tmpDate As Date

        oLustActivity.Retrieve(EventActivityID)
        If EventActivityID = 0 Then
            oLustActivity.FacilityID = oLocalLustEvent.FacilityID
        End If

        dtComments = oComments.GetComments("Technical", 23, oLustActivity.ActivityID).Tables(0)
        If dtComments.Rows.Count > 0 Then
            oComments.Retrieve(dtComments.Rows(0)("COMMENT_ID"), dtComments.Rows(0)("USER ID"))
            If oComments.Deleted = False Then
                txtComments.Text = oComments.Comments
            Else
                txtComments.Text = String.Empty
            End If
        Else
            txtComments.Text = String.Empty
        End If

        bolLoading = True

        PopulateLustActivities()
        If Mode = 0 Then ' Add Mode
            dtPickStartDate.Value = Now.Date
            'oLustActivity.Started = Now.Date
            gbGWSSamples.Visible = False
            oLustActivity.EventID = oLocalLustEvent.ID
            UIUtilsGen.SetDatePickerValue(dtPickCompletedDate, tmpDate)
            UIUtilsGen.SetDatePickerValue(dtPickCloseDate, tmpDate)
            UIUtilsGen.SetDatePickerValue(dt1stGWSBelow, tmpDate)
            UIUtilsGen.SetDatePickerValue(dt2ndGWSBelow, tmpDate)

            btnSave.Enabled = True
        Else 'Update Mode
            UIUtilsGen.SetDatePickerValue(dtPickStartDate, oLustActivity.Started)
            UIUtilsGen.SetDatePickerValue(dtPickCompletedDate, oLustActivity.Completed)
            UIUtilsGen.SetDatePickerValue(dtPickCloseDate, oLustActivity.Closed)
            UIUtilsGen.SetDatePickerValue(dt1stGWSBelow, oLustActivity.First_GWS_Below)
            UIUtilsGen.SetDatePickerValue(dt2ndGWSBelow, oLustActivity.Second_GWS_Below)
            btnSave.Enabled = False

            If oLustActivity.RemSystemID <= 0 AndAlso (cmbActivity.SelectedValue = 691 Or cmbActivity.SelectedValue = 692 Or cmbActivity.SelectedValue = 695 Or cmbActivity.SelectedValue = 1530 Or cmbActivity.SelectedValue = 699) Then
                Me.btAddSystem.Visible = True
            End If

        End If
        SetFormControls()

        bolLoading = False
    End Sub

    Private Sub dtPickStartDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickStartDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtPickStartDate)
        FillDateobjectValues(oLustActivity.Started, dtPickStartDate.Text)
    End Sub
    Private Sub dtPickCompletedDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickCompletedDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtPickCompletedDate)
        Me.FillDateobjectValues(oLustActivity.Completed, dtPickCompletedDate.Text)
    End Sub
    Private Sub dtPickCloseDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickCloseDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtPickCloseDate)
        Me.FillDateobjectValues(oLustActivity.Closed, dtPickCloseDate.Text)
    End Sub
    Private Sub dt1stGWSBelow_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dt1stGWSBelow.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dt1stGWSBelow)
        Me.FillDateobjectValues(oLustActivity.First_GWS_Below, dt1stGWSBelow.Text)
    End Sub
    Private Sub dt2ndGWSBelow_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dt2ndGWSBelow.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dt2ndGWSBelow)
        Me.FillDateobjectValues(oLustActivity.Second_GWS_Below, dt2ndGWSBelow.Text)
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub cmbActivity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbActivity.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLustActivity.Type = cmbActivity.SelectedValue
        ' 1.b Default Start date to current Date if Activity is one of the foll
        '  692 - REM Dual Phase
        '  691 - REM AS/SVE
        '  695 - REM Pump and Treat
        ' 1530 - REM - Dual Phase System Restart
        If cmbActivity.SelectedValue = 691 Or cmbActivity.SelectedValue = 692 Or cmbActivity.SelectedValue = 695 Or cmbActivity.SelectedValue = 1530 Then
            Me.dtPickStartDate.Checked = False
            UIUtilsGen.ToggleDateFormat(dtPickStartDate)
            FillDateobjectValues(oLustActivity.Started, dtPickStartDate.Text)
        Else
            dtPickStartDate.Checked = True
            dtPickStartDate.Value = Now.Date
            UIUtilsGen.ToggleDateFormat(dtPickStartDate)
            oLustActivity.Started = Now.Date
        End If
        SetFormControls()
    End Sub

    Private Sub SetFormControls()
        Dim oLustDocuments As New MUSTER.BusinessLogic.pLustEventDocument
        Dim colLustDocuments As New MUSTER.Info.LustDocumentCollection
        Dim oLustDocumentInfo As New MUSTER.Info.LustDocumentInfo
        Dim bolNonClosedDoc As Boolean = False
        Dim dt As Date
        ' --------------------------------------------------------------------------------
        'b.	If the added Activity is one of the following: 
        '         i.GWS()
        '        ii.	GWS –After System Shutdown 
        '       iii.	REM-AS/SVE 
        '        iv.	REM – Dual Phase 
        '         v.	REM Natural Attenuation
        '        vi.	REM-Pump & Treat System
        '        vii.   REM - Dual Phase System Restart
        'Then:
        '         i.	Display empty 1st GWS Below and 2nd GWS Below 
        '                   date fields on the Activity UI.  
        ' --------------------------------------------------------------------------------
        Try

            If cmbActivity.SelectedValue = 675 _
                Or cmbActivity.SelectedValue = 676 _
                Or cmbActivity.SelectedValue = 691 _
                Or cmbActivity.SelectedValue = 692 _
                Or cmbActivity.SelectedValue = 694 _
                Or cmbActivity.SelectedValue = 695 _
                Or cmbActivity.SelectedValue = 1530 _
            Then
                gbGWSSamples.Visible = True
            Else
                gbGWSSamples.Visible = False
            End If

            colLustDocuments = oLustDocuments.GetAllbyACTIVITYID(EventActivityID)

            ' --------------------------------------------------------------------------------
            ' If the TF Status is EUD or NTFE and the Activity is not NFA or NFA Under RBCA, 
            ' technically completed date is not available
            ' --------------------------------------------------------------------------------
            For Each oLustDocumentInfo In colLustDocuments.Values
                If oLustDocumentInfo.DocClosedDate = dt Then
                    bolUncompletedDoc = True
                End If
            Next
            If (TFStatus = 617 Or TFStatus = 620) And Not (cmbActivity.SelectedValue = 683 Or cmbActivity.SelectedValue = 684) Then
                dtPickCompletedDate.Enabled = False
            Else
                'If documents are open (no closed date) AND
                '(activity completed date is not null  and closed date is null) = false then disable the technically completed date picker
                'If bolUncompletedDoc And Not (oLustActivity.Completed <> dt And oLustActivity.Closed = dt) Then
                ' dtPickCompletedDate.Enabled = False
                'Else
                    dtPickCompletedDate.Enabled = True
                'End If
            End If

            ' --------------------------------------------------------------------------------
            '   If the Activity is NFA or NFA Under RBCA, closed date is not available
            ' --------------------------------------------------------------------------------
            If cmbActivity.SelectedValue = 683 _
            Or cmbActivity.SelectedValue = 684 Then
                dtPickCloseDate.Enabled = False
            Else
                For Each oLustDocumentInfo In colLustDocuments.Values
                    If oLustDocumentInfo.DocClosedDate = dt Then
                        bolNonClosedDoc = True
                    End If
                Next
                If bolNonClosedDoc Then
                    dtPickCloseDate.Enabled = False
                Else
                    dtPickCloseDate.Enabled = True
                End If
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception(ex.Message, ex))
            MyErr.ShowDialog()

        Finally

        End Try

    End Sub

#End Region

#Region " Populate Routines "

    Private Sub PopulateLustActivities()
        Try
            Dim dtLustActivity As DataTable = oLustActivity.PopulateLustActivities
            If Not IsNothing(dtLustActivity) Then
                cmbActivity.DataSource = dtLustActivity
                cmbActivity.DisplayMember = "PROPERTY_NAME"
                cmbActivity.ValueMember = "PROPERTY_ID"
            Else
                cmbActivity.DataSource = Nothing
            End If
            If Mode = 0 Then
                cmbActivity.SelectedIndex = -1
                cmbActivity.Enabled = True
            Else
                cmbActivity.SelectedValue = oLustActivity.Type
                cmbActivity.Enabled = False
            End If


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Lust Event Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub



    Private Sub FillDateobjectValues(ByRef currentObj As Object, ByVal value As String)

        If value.Length > 0 And value <> "__/__/____" Then
            currentObj = CType(value, Date)
        Else
            currentObj = "#12:00:00AM#"
        End If
    End Sub


#End Region
    Private Sub oLustActivity_LustEventChanged(ByVal bolValue As Boolean) Handles oLustActivity.LustEventChanged

        If bolValue = True Or txtComments.Text <> oComments.Comments Then
            btnSave.Enabled = True
        Else
            btnSave.Enabled = False
        End If
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        ProcessSaveEvent()
    End Sub
    Private Sub InsertComment()
        Dim oCommentInfo As New MUSTER.Info.CommentsInfo
        Try

            With oCommentInfo
                .CommentDate = Now.Date
                .Comments = txtComments.Text
                .CommentsScope = "External"
                .EntityID = oLustActivity.ActivityID
                .EntityType = 23
                .ModuleName = "Technical"
                .UserID = MusterContainer.AppUser.ID

                If .ID <= 0 Then
                    .CreatedBy = MusterContainer.AppUser.ID
                Else
                    .ModifiedBy = MusterContainer.AppUser.ID
                End If

            End With

            oComments.Add(oCommentInfo)
            oComments.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Insert Comment " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try

    End Sub
    Private Sub txtComments_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtComments.LostFocus

        If Mode <> 0 And oComments.ID > 0 Then 'If Not Add Mode and a Comment already exists
            If txtComments.Text = String.Empty Then
                oComments.Deleted = True
                btnSave.Enabled = True
            ElseIf txtComments.Text <> oComments.Comments Then
                oComments.Deleted = False
                oComments.Comments = txtComments.Text
                btnSave.Enabled = True
            End If
        Else
            If txtComments.Text <> String.Empty Or oLustActivity.IsDirty Then
                btnSave.Enabled = True
            Else
                btnSave.Enabled = False
            End If
        End If
    End Sub
    Private Sub txtComments_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtComments.TextChanged
        btnSave.Enabled = True
    End Sub

    Private Sub AddRemediationSystem()
        Dim frmRemSysList As New RemediationSystemList
        Try

            frmRemSysList.CallingForm = Me
            frmRemSysList.Mode = 0 ' Add
            frmRemSysList.EventActivityID = oLustActivity.ActivityID

            frmRemSysList.ShowDialog()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Add Remediation System" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            frmRemSysList = Nothing
        End Try
    End Sub

    Private Sub Activity_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

    End Sub

    Private Sub Activity_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If oLustActivity.IsDirty Then
            If MsgBox("Do you wish to save changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                ProcessSaveEvent()
            End If
        End If
    End Sub

    Private Sub ProcessSaveEvent()
        Dim bolAllClosed As Boolean
        Try
            If cmbActivity.Text = "" Then
                MsgBox("Please select an Activity")
                Exit Sub
            End If
            If oLustActivity.Closed <> "01/01/0001" And oLustActivity.Closed < oLustActivity.Started Then
                MsgBox("Closed Date Cannot Be Before Start Date")
                Exit Sub
            End If
            If oLustActivity.Completed <> "01/01/0001" And oLustActivity.Completed < oLustActivity.Started Then
                MsgBox("Completed Date Cannot Be Before Start Date")
                Exit Sub
            End If

            If Not (cmbActivity.SelectedValue = 691 Or cmbActivity.SelectedValue = 692 Or cmbActivity.SelectedValue = 695 Or cmbActivity.SelectedValue = 1530) Then
                If oLustActivity.Started = "01/01/0001" Then
                    MsgBox("Start Date Required")
                    Exit Sub
                End If
            End If

            If gbGWSSamples.Visible Then
                If oLustActivity.First_GWS_Below < oLustActivity.Started And oLustActivity.First_GWS_Below <> "01/01/0001" And oLustActivity.Started <> "01/01/0001" Then
                    MsgBox("1st GWS Below Date Cannot Be Before Start Date")
                    Exit Sub
                End If
                If oLustActivity.Second_GWS_Below < oLustActivity.First_GWS_Below And oLustActivity.First_GWS_Below <> "01/01/0001" And oLustActivity.Second_GWS_Below <> "01/01/0001" Then
                    MsgBox("2nd GWS Below Date Cannot Be Before 1st GWS Below Date")
                    Exit Sub
                End If
                If oLustActivity.Second_GWS_Below < oLustActivity.Started And oLustActivity.Second_GWS_Below <> "01/01/0001" And oLustActivity.Started <> "01/01/0001" Then
                    MsgBox("2nd GWS Below Date Cannot Be Before Start Date")
                    Exit Sub
                End If
                If oLustActivity.Completed < oLustActivity.First_GWS_Below And oLustActivity.First_GWS_Below <> "01/01/0001" And oLustActivity.Completed <> "01/01/0001" Then
                    MsgBox("Completed Date Cannot Be Before 1st GWS Below Date")
                    Exit Sub
                End If
                If oLustActivity.Completed < oLustActivity.Second_GWS_Below And oLustActivity.Second_GWS_Below <> "01/01/0001" And oLustActivity.Completed <> "01/01/0001" Then
                    MsgBox("Completed Date Cannot Be Before 2nd GWS Below Date")
                    Exit Sub
                End If
                If oLustActivity.Closed < oLustActivity.First_GWS_Below And oLustActivity.First_GWS_Below <> "01/01/0001" And oLustActivity.Closed <> "01/01/0001" Then
                    MsgBox("Closed Date Cannot Be Before 1st GWS Below Date")
                    Exit Sub
                End If
                If oLustActivity.Closed < oLustActivity.Second_GWS_Below And oLustActivity.Second_GWS_Below <> "01/01/0001" And oLustActivity.Closed <> "01/01/0001" Then
                    MsgBox("Closed Date Cannot Be Before 2nd GWS Below Date")
                    Exit Sub
                End If
            End If

            If oLustActivity.IsDirty Then
                If oLustActivity.ActivityID <= 0 Then
                    oLustActivity.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oLustActivity.ModifiedBy = MusterContainer.AppUser.ID
                End If

                oLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal, bolUncompletedDoc)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            End If

            If Mode = 0 Then
                If txtComments.Text <> String.Empty Then
                    InsertComment()
                End If
            Else
                If oComments.ID > 0 Then
                    If oComments.IsDirty Then
                        oComments.ModifiedBy = MusterContainer.AppUser.ID
                        oComments.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    End If
                ElseIf txtComments.Text <> String.Empty Then
                    InsertComment()
                End If
            End If
            If cmbActivity.SelectedValue = 699 Then
                MsgBox("Have you requested a wastewater permit?")
            End If

            '   If Mode = 0 And (cmbActivity.SelectedValue = 691 Or cmbActivity.SelectedValue = 692 Or cmbActivity.SelectedValue = 695 Or cmbActivity.SelectedValue = 1530 Or cmbActivity.SelectedValue = 699) Then
            'removed Turnkey Process
            If Mode = 0 And (cmbActivity.SelectedValue = 691 Or cmbActivity.SelectedValue = 692 Or cmbActivity.SelectedValue = 695 Or cmbActivity.SelectedValue = 1530) Then

                If MsgBox("Would you like to tie this Activity to a remediation system?", MsgBoxStyle.YesNo, "REM-Dual Phase System Setup") = MsgBoxResult.Yes Then

                    AddRemediationSystem()
                    oLustActivity.AgeThreshold = 0
                    oLustActivity.Retrieve(oLustActivity.ActivityID)
                    If Not (oLustActivity.RemSystemID > 0) Then
                        If MsgBox("No remediation system was assigned.  Do you want to do this now?" & vbCrLf & "You will be prompted again about adding a remediation system.", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            ProcessSaveEvent()
                            Exit Sub
                        Else
                            oLustActivity.Deleted = True
                            oLustActivity.ModifiedBy = MusterContainer.AppUser.ID
                            oLustActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If


            If oLocalLustEvent.OpenActivities(oLocalLustEvent.ID) = 0 Then
                EventActivityID = 0
                Mode = 0
                ProcessLoad()
            Else
                Me.Close()
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Save Activity " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        Finally

        End Try
    End Sub

    Private Sub btnAddSystem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAddSystem.Click

        AddRemediationSystem()
        oLustActivity.AgeThreshold = 0
        oLustActivity.Retrieve(oLustActivity.ActivityID)

        If oLustActivity.RemSystemID > 0 Then

            btAddSystem.Visible = False
            MsgBox("Remediation System has been tied to this activity")

        Else
            MsgBox("Adding of Remediation System to Activity canceled")
        End If

    End Sub
End Class

