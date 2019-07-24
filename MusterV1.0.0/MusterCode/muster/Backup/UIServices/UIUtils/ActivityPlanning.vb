Public Class ActivityPlanning
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

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
    Friend WithEvents pnlEditActivityPlan As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cboActivty As System.Windows.Forms.ComboBox
    Friend WithEvents lblActivityName As System.Windows.Forms.Label
    Friend WithEvents lblDuration As System.Windows.Forms.Label
    Friend WithEvents txtDuration As System.Windows.Forms.TextBox
    Friend WithEvents lblActivityPlanner As System.Windows.Forms.Label
    Friend WithEvents ugPlannedActivites As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblCost As System.Windows.Forms.Label
    Friend WithEvents lblSiteIDDesc As System.Windows.Forms.Label
    Friend WithEvents lblSiteNameDesc As System.Windows.Forms.Label
    Friend WithEvents lblSiteID As System.Windows.Forms.Label
    Friend WithEvents lblSiteName As System.Windows.Forms.Label
    Friend WithEvents lblEvent As System.Windows.Forms.Label
    Friend WithEvents LblEventDesc As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents BtnAdd As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents txtCost As System.Windows.Forms.TextBox
    Friend WithEvents btncancelEdit As System.Windows.Forms.Button
    Friend WithEvents btnSaveGoOn As System.Windows.Forms.Button
    Friend WithEvents btnCalc As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlEditActivityPlan = New System.Windows.Forms.Panel
        Me.btnCalc = New System.Windows.Forms.Button
        Me.btnSaveGoOn = New System.Windows.Forms.Button
        Me.btncancelEdit = New System.Windows.Forms.Button
        Me.txtCost = New System.Windows.Forms.TextBox
        Me.lblCost = New System.Windows.Forms.Label
        Me.txtDuration = New System.Windows.Forms.TextBox
        Me.lblDuration = New System.Windows.Forms.Label
        Me.lblActivityName = New System.Windows.Forms.Label
        Me.cboActivty = New System.Windows.Forms.ComboBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.lblActivityPlanner = New System.Windows.Forms.Label
        Me.ugPlannedActivites = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lblSiteIDDesc = New System.Windows.Forms.Label
        Me.lblSiteNameDesc = New System.Windows.Forms.Label
        Me.lblSiteID = New System.Windows.Forms.Label
        Me.lblSiteName = New System.Windows.Forms.Label
        Me.lblEvent = New System.Windows.Forms.Label
        Me.LblEventDesc = New System.Windows.Forms.Label
        Me.BtnAdd = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.pnlEditActivityPlan.SuspendLayout()
        CType(Me.ugPlannedActivites, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlEditActivityPlan
        '
        Me.pnlEditActivityPlan.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlEditActivityPlan.Controls.Add(Me.btnCalc)
        Me.pnlEditActivityPlan.Controls.Add(Me.btnSaveGoOn)
        Me.pnlEditActivityPlan.Controls.Add(Me.btncancelEdit)
        Me.pnlEditActivityPlan.Controls.Add(Me.txtCost)
        Me.pnlEditActivityPlan.Controls.Add(Me.lblCost)
        Me.pnlEditActivityPlan.Controls.Add(Me.txtDuration)
        Me.pnlEditActivityPlan.Controls.Add(Me.lblDuration)
        Me.pnlEditActivityPlan.Controls.Add(Me.lblActivityName)
        Me.pnlEditActivityPlan.Controls.Add(Me.cboActivty)
        Me.pnlEditActivityPlan.Controls.Add(Me.btnSave)
        Me.pnlEditActivityPlan.Enabled = False
        Me.pnlEditActivityPlan.Location = New System.Drawing.Point(8, 248)
        Me.pnlEditActivityPlan.Name = "pnlEditActivityPlan"
        Me.pnlEditActivityPlan.Size = New System.Drawing.Size(568, 80)
        Me.pnlEditActivityPlan.TabIndex = 0
        '
        'btnCalc
        '
        Me.btnCalc.Location = New System.Drawing.Point(200, 48)
        Me.btnCalc.Name = "btnCalc"
        Me.btnCalc.Size = New System.Drawing.Size(22, 21)
        Me.btnCalc.TabIndex = 47
        Me.btnCalc.Text = "C"
        '
        'btnSaveGoOn
        '
        Me.btnSaveGoOn.Location = New System.Drawing.Point(440, 48)
        Me.btnSaveGoOn.Name = "btnSaveGoOn"
        Me.btnSaveGoOn.Size = New System.Drawing.Size(120, 23)
        Me.btnSaveGoOn.TabIndex = 46
        Me.btnSaveGoOn.Text = "Save && Go to Next"
        '
        'btncancelEdit
        '
        Me.btncancelEdit.BackColor = System.Drawing.Color.Brown
        Me.btncancelEdit.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btncancelEdit.Location = New System.Drawing.Point(536, 0)
        Me.btncancelEdit.Name = "btncancelEdit"
        Me.btncancelEdit.Size = New System.Drawing.Size(24, 16)
        Me.btncancelEdit.TabIndex = 45
        Me.btncancelEdit.Text = "X"
        '
        'txtCost
        '
        Me.txtCost.Enabled = False
        Me.txtCost.Location = New System.Drawing.Point(80, 48)
        Me.txtCost.Name = "txtCost"
        Me.txtCost.Size = New System.Drawing.Size(120, 20)
        Me.txtCost.TabIndex = 5
        Me.txtCost.Text = ""
        '
        'lblCost
        '
        Me.lblCost.Location = New System.Drawing.Point(-16, 48)
        Me.lblCost.Name = "lblCost"
        Me.lblCost.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCost.Size = New System.Drawing.Size(96, 16)
        Me.lblCost.TabIndex = 4
        Me.lblCost.Text = "Estimated Cost :"
        Me.lblCost.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDuration
        '
        Me.txtDuration.Enabled = False
        Me.txtDuration.Location = New System.Drawing.Point(432, 16)
        Me.txtDuration.Name = "txtDuration"
        Me.txtDuration.Size = New System.Drawing.Size(120, 20)
        Me.txtDuration.TabIndex = 3
        Me.txtDuration.Text = ""
        '
        'lblDuration
        '
        Me.lblDuration.Location = New System.Drawing.Point(312, 16)
        Me.lblDuration.Name = "lblDuration"
        Me.lblDuration.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDuration.Size = New System.Drawing.Size(112, 16)
        Me.lblDuration.TabIndex = 2
        Me.lblDuration.Text = "Duration in Months :"
        Me.lblDuration.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblActivityName
        '
        Me.lblActivityName.Location = New System.Drawing.Point(16, 16)
        Me.lblActivityName.Name = "lblActivityName"
        Me.lblActivityName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblActivityName.Size = New System.Drawing.Size(56, 16)
        Me.lblActivityName.TabIndex = 1
        Me.lblActivityName.Text = "Activity :"
        Me.lblActivityName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboActivty
        '
        Me.cboActivty.Enabled = False
        Me.cboActivty.Location = New System.Drawing.Point(80, 16)
        Me.cboActivty.Name = "cboActivty"
        Me.cboActivty.Size = New System.Drawing.Size(216, 21)
        Me.cboActivty.TabIndex = 0
        Me.cboActivty.Text = "ComboBox1"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(320, 48)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(112, 23)
        Me.btnSave.TabIndex = 44
        Me.btnSave.Text = "Save Changes"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(504, 336)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        '
        'lblActivityPlanner
        '
        Me.lblActivityPlanner.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblActivityPlanner.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblActivityPlanner.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblActivityPlanner.Location = New System.Drawing.Point(16, 80)
        Me.lblActivityPlanner.Name = "lblActivityPlanner"
        Me.lblActivityPlanner.Size = New System.Drawing.Size(560, 40)
        Me.lblActivityPlanner.TabIndex = 3
        Me.lblActivityPlanner.Text = "Planned Activities"
        Me.lblActivityPlanner.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ugPlannedActivites
        '
        Me.ugPlannedActivites.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPlannedActivites.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugPlannedActivites.DisplayLayout.Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
        Me.ugPlannedActivites.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.None
        Me.ugPlannedActivites.DisplayLayout.Override.AllowColSwapping = Infragistics.Win.UltraWinGrid.AllowColSwapping.NotAllowed
        Me.ugPlannedActivites.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugPlannedActivites.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False
        Me.ugPlannedActivites.DisplayLayout.Override.AllowGroupMoving = Infragistics.Win.UltraWinGrid.AllowGroupMoving.NotAllowed
        Me.ugPlannedActivites.DisplayLayout.Override.AllowGroupSwapping = Infragistics.Win.UltraWinGrid.AllowGroupSwapping.NotAllowed
        Me.ugPlannedActivites.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.False
        Me.ugPlannedActivites.DisplayLayout.Override.AllowRowLayoutCellSizing = Infragistics.Win.UltraWinGrid.RowLayoutSizing.None
        Me.ugPlannedActivites.DisplayLayout.Override.AllowRowLayoutLabelSizing = Infragistics.Win.UltraWinGrid.RowLayoutSizing.None
        Me.ugPlannedActivites.DisplayLayout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.False
        Me.ugPlannedActivites.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugPlannedActivites.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugPlannedActivites.Location = New System.Drawing.Point(16, 120)
        Me.ugPlannedActivites.Name = "ugPlannedActivites"
        Me.ugPlannedActivites.Size = New System.Drawing.Size(560, 120)
        Me.ugPlannedActivites.TabIndex = 35
        Me.ugPlannedActivites.Text = "Activities"
        '
        'lblSiteIDDesc
        '
        Me.lblSiteIDDesc.Location = New System.Drawing.Point(16, 8)
        Me.lblSiteIDDesc.Name = "lblSiteIDDesc"
        Me.lblSiteIDDesc.Size = New System.Drawing.Size(64, 23)
        Me.lblSiteIDDesc.TabIndex = 36
        Me.lblSiteIDDesc.Text = "Site ID:"
        Me.lblSiteIDDesc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSiteNameDesc
        '
        Me.lblSiteNameDesc.Location = New System.Drawing.Point(0, 40)
        Me.lblSiteNameDesc.Name = "lblSiteNameDesc"
        Me.lblSiteNameDesc.Size = New System.Drawing.Size(80, 23)
        Me.lblSiteNameDesc.TabIndex = 37
        Me.lblSiteNameDesc.Text = "Site Name :"
        Me.lblSiteNameDesc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSiteID
        '
        Me.lblSiteID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSiteID.Location = New System.Drawing.Point(80, 8)
        Me.lblSiteID.Name = "lblSiteID"
        Me.lblSiteID.Size = New System.Drawing.Size(88, 24)
        Me.lblSiteID.TabIndex = 38
        Me.lblSiteID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblSiteName
        '
        Me.lblSiteName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSiteName.Location = New System.Drawing.Point(80, 40)
        Me.lblSiteName.Name = "lblSiteName"
        Me.lblSiteName.Size = New System.Drawing.Size(200, 24)
        Me.lblSiteName.TabIndex = 39
        Me.lblSiteName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEvent
        '
        Me.lblEvent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblEvent.Location = New System.Drawing.Point(480, 8)
        Me.lblEvent.Name = "lblEvent"
        Me.lblEvent.Size = New System.Drawing.Size(96, 24)
        Me.lblEvent.TabIndex = 41
        Me.lblEvent.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LblEventDesc
        '
        Me.LblEventDesc.Location = New System.Drawing.Point(400, 8)
        Me.LblEventDesc.Name = "LblEventDesc"
        Me.LblEventDesc.Size = New System.Drawing.Size(80, 23)
        Me.LblEventDesc.TabIndex = 40
        Me.LblEventDesc.Text = "Tec Event #:"
        Me.LblEventDesc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BtnAdd
        '
        Me.BtnAdd.Enabled = False
        Me.BtnAdd.Location = New System.Drawing.Point(8, 336)
        Me.BtnAdd.Name = "BtnAdd"
        Me.BtnAdd.TabIndex = 42
        Me.BtnAdd.Text = "Add New"
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(96, 336)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.TabIndex = 43
        Me.btnDelete.Text = "Delete"
        '
        'ActivityPlanning
        '
        Me.AcceptButton = Me.btnCalc
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 373)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.BtnAdd)
        Me.Controls.Add(Me.lblEvent)
        Me.Controls.Add(Me.LblEventDesc)
        Me.Controls.Add(Me.lblSiteName)
        Me.Controls.Add(Me.lblSiteID)
        Me.Controls.Add(Me.lblSiteNameDesc)
        Me.Controls.Add(Me.lblSiteIDDesc)
        Me.Controls.Add(Me.ugPlannedActivites)
        Me.Controls.Add(Me.lblActivityPlanner)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.pnlEditActivityPlan)
        Me.MaximumSize = New System.Drawing.Size(600, 400)
        Me.MinimumSize = New System.Drawing.Size(600, 400)
        Me.Name = "ActivityPlanning"
        Me.Text = "Future Activity Planner"
        Me.pnlEditActivityPlan.ResumeLayout(False)
        CType(Me.ugPlannedActivites, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "private members"

    Private nFacilityID As Integer
    Private strFacility As String
    Private nEventID As Integer
    Private nEventSeq As Integer
    Private nLastIndex As Integer = -1
    Private nLastNewStatus As Boolean = False
    Private nModuleID As Integer
    Private boAct As BusinessLogic.pTecFinancialActivityPlanner
    Private bCancelClose As Boolean = False


#End Region

#Region "Public Properties"
    Public Property ActvityBO() As BusinessLogic.pTecFinancialActivityPlanner

        Get

            If boAct Is Nothing Then
                boAct = New BusinessLogic.pTecFinancialActivityPlanner
            End If

            Return boAct
        End Get

        Set(ByVal Value As BusinessLogic.pTecFinancialActivityPlanner)
            boAct = Value
        End Set

    End Property

#End Region
#Region "Construct"

    Sub New(ByVal moduleID As Integer, ByVal facility_ID As Integer, ByVal facility As String, ByVal event_ID As Integer, ByVal event_Sequence As Integer)

        Me.New()
        nFacilityID = facility_ID
        nEventID = event_ID
        nEventSeq = event_Sequence
        strFacility = facility
        nModuleID = moduleID

    End Sub

#End Region

#Region "private methods"

    Sub LoadActivityCombo()



        cboActivty.DataSource = ActvityBO.PopulateTecActivityListForCosts(nEventID)
        cboActivty.ValueMember = "PROPERTY_ID"
        cboActivty.DisplayMember = "PROPERTY_NAME"
    End Sub


    Sub LoadEventDataAndRights()


        lblSiteName.Text = strFacility
        lblSiteID.Text = nFacilityID
        lblEvent.Text = nEventSeq

        If nModuleID = 614 Then



            BtnAdd.Enabled = True
            btnDelete.Enabled = True
            txtDuration.Enabled = True
            Text = "Technical Planner for Activities"


        ElseIf nModuleID = 616 Then
            txtCost.Enabled = True
            btnCalc.Visible = False
            Text = "Financial Planner for Technical Activities "

        End If

    End Sub

    Sub CleanDataEntryBoxes()

        cboActivty.SelectedIndex = -1
        cboActivty.DropDownStyle = ComboBoxStyle.DropDownList
        cboActivty.Enabled = False

        txtCost.Text = String.Empty
        txtDuration.Text = String.Empty

    End Sub


    Sub EditActivityRecord(ByVal actID As Integer, ByVal isNew As Boolean, ByVal delete As Boolean)

        nLastNewStatus = isNew

        txtCost.DataBindings.Clear()
        txtDuration.DataBindings.Clear()
        cboActivty.DataBindings.Clear()

        btnSaveGoOn.Enabled = True
        cboActivty.Enabled = False




        If isNew Then
            ActvityBO = New BusinessLogic.pTecFinancialActivityPlanner

            If nModuleID = 614 Then
                cboActivty.Enabled = True
            End If


            With ActvityBO
                .ActivityTypeID = 0
                .ID = nEventID

            End With

            btnSaveGoOn.Text = "Save && Add New"

        Else

            ActvityBO.Retrieve(nEventID, actID)

            If delete Then

                ActvityBO.Duration = -1
                SaveChanges(True)
                Exit Sub
            End If

            btnSaveGoOn.Text = "Save && Go to Next"

            If ugPlannedActivites.ActiveRow.Index = ugPlannedActivites.Rows.Count - 1 Then
                btnSaveGoOn.Enabled = False
            End If

            cboActivty.Enabled = False


        End If


        txtCost.DataBindings.Add("Text", ActvityBO, "Cost")
        txtDuration.DataBindings.Add("Text", ActvityBO, "Duration")
        cboActivty.DataBindings.Add("SelectedValue", ActvityBO, "ActivityTypeID")

        pnlEditActivityPlan.Enabled = True

        If nModuleID = 614 AndAlso cboActivty.Enabled Then
            cboActivty.Focus()
        ElseIf nModuleID = 614 Then
            txtDuration.Focus()
        Else
            txtCost.Focus()
        End If
    End Sub

    Sub SaveChanges(Optional ByVal delete As Boolean = False)

        Dim retValue As String

        Try
            If delete Or ((Me.ActvityBO.ID > 0 Or Me.ActvityBO.ID < -9999) AndAlso IIf(nModuleID = 614, IsNumeric(txtDuration.Text) AndAlso Convert.ToInt32(Me.txtDuration.Text) > 0, _
                                                          IsNumeric(txtCost.Text) AndAlso Convert.ToInt32(Me.txtCost.Text) > 0)) Then

                If txtCost.Text.Length > 0 AndAlso IsNumeric(txtCost.Text) Then
                    ActvityBO.Cost = Convert.ToDouble(txtCost.Text)
                Else
                    ActvityBO.Cost = 0
                End If

                ActvityBO.Save(nModuleID, MusterContainer.AppUser.UserKey, retValue)

                If Not ugPlannedActivites.ActiveRow Is Nothing Then
                    nLastIndex = Me.ugPlannedActivites.ActiveRow.Index
                End If

            Else
                retValue = "Please make sure you select an Activity and that all fields are numeric and greater than 0 before you save."
            End If


            If retValue.Length > 0 Then
                MsgBox(retValue)
                Exit Sub
            End If

            txtCost.DataBindings.Clear()
            txtDuration.DataBindings.Clear()
            cboActivty.DataBindings.Clear()

            pnlEditActivityPlan.Enabled = False
            LoadActivityCombo()
            LoadActivityList()


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Sub LoadActivityList()
        ugPlannedActivites.DataSource = Me.ActvityBO.PopulateActivityPlanning(nEventID)
    End Sub

    Sub InitializeForm()

        LoadActivityCombo()
        LoadEventDataAndRights()
        LoadActivityList()
        CleanDataEntryBoxes()




    End Sub

#End Region

#Region "public Form Events"


    Private Sub ugActivities_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugPlannedActivites.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If

            EditActivityRecord(ugPlannedActivites.ActiveRow.Cells("Activity Type").Value, False, False)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Public Sub LoadForm(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        InitializeForm()
    End Sub

    Public Sub CloseForm(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Closed
        If Not boAct Is Nothing Then
            boAct = Nothing
        End If
    End Sub

    Public Sub btnAddClicked(ByVal sender As Object, ByVal e As EventArgs) Handles BtnAdd.Click
        EditActivityRecord(-1, True, False)
    End Sub

    Public Sub btnDeleteClicked(ByVal sender As Object, ByVal e As EventArgs) Handles btnDelete.Click

        If Not Me.ugPlannedActivites.ActiveRow Is Nothing Then

            EditActivityRecord(ugPlannedActivites.ActiveRow.Cells("Activity Type").Value, False, True)
        Else

            If Me.ugPlannedActivites.Rows.Count > 0 Then
                MsgBox("Please Select a record to delete first")
            Else
                MsgBox("you have no planned activities to delete")
            End If
        End If

    End Sub

    Public Sub btnSaveClicked(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        SaveChanges()
    End Sub

    Public Sub btnSaveGoOnClicked(ByVal sender As Object, ByVal e As EventArgs) Handles btnSaveGoOn.Click

        SaveChanges()



        With ugPlannedActivites

            If Not nLastNewStatus AndAlso nLastIndex < (.Rows.Count - 1) Then
                .ActiveRow = .Rows(nLastIndex + 1)
                EditActivityRecord(ugPlannedActivites.ActiveRow.Cells("Activity Type").Value, False, False)
            ElseIf nLastNewStatus Then
                EditActivityRecord(-1, True, False)
            End If

        End With

    End Sub

    Public Sub BtnCloseClicked(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click

        pnlEditActivityPlan.Enabled = False
        Close()
    End Sub

    Private Sub CalculateCost()

        If nModuleID = 614 AndAlso pnlEditActivityPlan.Enabled AndAlso txtDuration.Text.Length > 0 _
          AndAlso IsNumeric(txtDuration.Text) Then

            Dim newAct As New BusinessLogic.pTecAct

            newAct.Retrieve(ActvityBO.ActivityTypeID)

            Select Case newAct.CostMode

                Case Info.TecActInfo.ActivityCostModeEnum.PerActivity
                    txtCost.Text = newAct.Cost.ToString
                Case Info.TecActInfo.ActivityCostModeEnum.PerMonth
                    txtCost.Text = (newAct.Cost * ActvityBO.Duration).ToString
                Case Info.TecActInfo.ActivityCostModeEnum.PerQuarter
                    txtCost.Text = (newAct.Cost * (ActvityBO.Duration / 3.0)).ToString
                Case Info.TecActInfo.ActivityCostModeEnum.PerYear
                    txtCost.Text = (newAct.Cost * (ActvityBO.Duration / 12)).ToString

            End Select

            newAct = Nothing

        ElseIf pnlEditActivityPlan.Enabled AndAlso (txtDuration.Text.Length = 0 _
          OrElse Not IsNumeric(txtDuration.Text)) Then
            ActvityBO.Cost = 0.0

        End If

    End Sub

    Sub changeActivities(ByVal sender As Object, ByVal e As EventArgs) Handles cboActivty.SelectedIndexChanged
        CalculateCost()
    End Sub

    Sub changeDuration(ByVal sender As Object, ByVal e As EventArgs) Handles txtDuration.TextChanged
        CalculateCost()
    End Sub

    Private Sub btnCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalc.Click
        CalculateCost()
    End Sub

    Private Sub btncancelEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancelEdit.Click
        pnlEditActivityPlan.Enabled = False
    End Sub


    Private Sub pnlEditActivityPlan_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlEditActivityPlan.EnabledChanged

        If Not Me.pnlEditActivityPlan.Enabled AndAlso ActvityBO.IsDirty Then

            If MsgBox("You have unsaved changes.  Are you sure you to exit?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then

                Me.pnlEditActivityPlan.Enabled = True
                bCancelClose = True
                Exit Sub

            End If

            ActvityBO = Nothing

        End If

        If Not Me.pnlEditActivityPlan.Enabled Then
            CleanDataEntryBoxes()
        End If


        If nModuleID <> 616 Then
            BtnAdd.Enabled = Not pnlEditActivityPlan.Enabled
        End If


    End Sub

    Private Sub ActivityPlanning_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If bCancelClose Then
            bCancelClose = False
            e.Cancel = True
        End If

    End Sub
#End Region


#Region "Grid Setup"

    Private Sub ugPlannedActivites_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugPlannedActivites.InitializeLayout

        With e.Layout.Bands(0)

            e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            .Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            .Columns("Event ID").Hidden = True
            .Columns("Facility").Hidden = True
            .Columns("Event").Hidden = True
            .Columns("Activity Type").Hidden = True

            .Columns("Activity").Width = 65
            .Columns("Duration").Width = 20
            .Columns("Cost").Width = 15

            With e.Layout.Bands(0)

                .Columns("Activity").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                .Columns("Duration").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                .Columns("Cost").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            End With
            e.Layout.AutoFitColumns = True

            .Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.None



        End With

    End Sub

    Private Sub cboActivty_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboActivty.SelectedIndexChanged

        txtDuration.Focus()
    End Sub

#End Region


End Class
