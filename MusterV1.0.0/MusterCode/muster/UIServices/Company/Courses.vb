Public Class Courses
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    Private WithEvents oCourse As New MUSTER.BusinessLogic.pCourse
    Private WithEvents oProvider As New MUSTER.BusinessLogic.pProvider
    Private oCourseInfo As MUSTER.Info.CourseInfo
    Private bolValidationFlg As Boolean = False
    Private nCourseID As Integer = 0
    Private bolLoading As Boolean = False
    Private dtCourses As New DataTable
    Private nID As Integer = -1
    Dim returnVal As String = String.Empty

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef ParentForm As Windows.Forms.Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        If Not ParentForm Is Nothing Then
            Me.MdiParent = ParentForm
        End If
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
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents txtCourseTitle As System.Windows.Forms.TextBox
    Friend WithEvents lblCourseTitle As System.Windows.Forms.Label
    Friend WithEvents txtCourseDates As System.Windows.Forms.TextBox
    Friend WithEvents lblCourseDates As System.Windows.Forms.Label
    Friend WithEvents txtLocation As System.Windows.Forms.TextBox
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblProvider As System.Windows.Forms.Label
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents cmbProvider As System.Windows.Forms.ComboBox
    Friend WithEvents txtHours As System.Windows.Forms.TextBox
    Friend WithEvents lblHours As System.Windows.Forms.Label
    Friend WithEvents cmbCourseTitle As System.Windows.Forms.ComboBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents ugCourseDates As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.txtCourseTitle = New System.Windows.Forms.TextBox
        Me.lblCourseTitle = New System.Windows.Forms.Label
        Me.txtCourseDates = New System.Windows.Forms.TextBox
        Me.lblCourseDates = New System.Windows.Forms.Label
        Me.txtLocation = New System.Windows.Forms.TextBox
        Me.lblLocation = New System.Windows.Forms.Label
        Me.lblProvider = New System.Windows.Forms.Label
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.lblType = New System.Windows.Forms.Label
        Me.cmbProvider = New System.Windows.Forms.ComboBox
        Me.txtHours = New System.Windows.Forms.TextBox
        Me.lblHours = New System.Windows.Forms.Label
        Me.cmbCourseTitle = New System.Windows.Forms.ComboBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.ugCourseDates = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.ugCourseDates, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(280, 312)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 23)
        Me.btnDelete.TabIndex = 11
        Me.btnDelete.Text = "Delete"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(368, 312)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(192, 312)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 23)
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "Save"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(104, 312)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 23)
        Me.btnNew.TabIndex = 8
        Me.btnNew.Text = "New"
        '
        'chkActive
        '
        Me.chkActive.Checked = True
        Me.chkActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkActive.Location = New System.Drawing.Point(440, 40)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkActive.Size = New System.Drawing.Size(72, 16)
        Me.chkActive.TabIndex = 7
        Me.chkActive.Tag = "644"
        Me.chkActive.Text = " :Active"
        Me.chkActive.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'txtCourseTitle
        '
        Me.txtCourseTitle.Location = New System.Drawing.Point(112, 64)
        Me.txtCourseTitle.Multiline = True
        Me.txtCourseTitle.Name = "txtCourseTitle"
        Me.txtCourseTitle.Size = New System.Drawing.Size(296, 48)
        Me.txtCourseTitle.TabIndex = 0
        Me.txtCourseTitle.Text = ""
        '
        'lblCourseTitle
        '
        Me.lblCourseTitle.Location = New System.Drawing.Point(40, 48)
        Me.lblCourseTitle.Name = "lblCourseTitle"
        Me.lblCourseTitle.Size = New System.Drawing.Size(64, 24)
        Me.lblCourseTitle.TabIndex = 244
        Me.lblCourseTitle.Text = "Course Title:"
        Me.lblCourseTitle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtCourseDates
        '
        Me.txtCourseDates.Location = New System.Drawing.Point(616, 72)
        Me.txtCourseDates.Name = "txtCourseDates"
        Me.txtCourseDates.Size = New System.Drawing.Size(296, 20)
        Me.txtCourseDates.TabIndex = 3
        Me.txtCourseDates.Text = ""
        Me.txtCourseDates.Visible = False
        '
        'lblCourseDates
        '
        Me.lblCourseDates.Location = New System.Drawing.Point(528, 72)
        Me.lblCourseDates.Name = "lblCourseDates"
        Me.lblCourseDates.Size = New System.Drawing.Size(80, 17)
        Me.lblCourseDates.TabIndex = 242
        Me.lblCourseDates.Text = "Course Dates:"
        Me.lblCourseDates.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblCourseDates.Visible = False
        '
        'txtLocation
        '
        Me.txtLocation.Location = New System.Drawing.Point(616, 96)
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.Size = New System.Drawing.Size(200, 20)
        Me.txtLocation.TabIndex = 4
        Me.txtLocation.Text = ""
        Me.txtLocation.Visible = False
        '
        'lblLocation
        '
        Me.lblLocation.Location = New System.Drawing.Point(536, 96)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(72, 17)
        Me.lblLocation.TabIndex = 240
        Me.lblLocation.Text = "Location:"
        Me.lblLocation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblLocation.Visible = False
        '
        'lblProvider
        '
        Me.lblProvider.Location = New System.Drawing.Point(48, 114)
        Me.lblProvider.Name = "lblProvider"
        Me.lblProvider.Size = New System.Drawing.Size(56, 17)
        Me.lblProvider.TabIndex = 250
        Me.lblProvider.Text = "Provider:"
        Me.lblProvider.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbType
        '
        Me.cmbType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbType.DropDownWidth = 180
        Me.cmbType.ItemHeight = 13
        Me.cmbType.Location = New System.Drawing.Point(616, 120)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(144, 21)
        Me.cmbType.TabIndex = 5
        Me.cmbType.Visible = False
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(560, 120)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(48, 17)
        Me.lblType.TabIndex = 253
        Me.lblType.Text = "Type:"
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblType.Visible = False
        '
        'cmbProvider
        '
        Me.cmbProvider.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbProvider.DropDownWidth = 180
        Me.cmbProvider.ItemHeight = 13
        Me.cmbProvider.Location = New System.Drawing.Point(112, 114)
        Me.cmbProvider.Name = "cmbProvider"
        Me.cmbProvider.Size = New System.Drawing.Size(296, 21)
        Me.cmbProvider.TabIndex = 2
        '
        'txtHours
        '
        Me.txtHours.Location = New System.Drawing.Point(616, 144)
        Me.txtHours.Name = "txtHours"
        Me.txtHours.Size = New System.Drawing.Size(72, 20)
        Me.txtHours.TabIndex = 6
        Me.txtHours.Text = ""
        Me.txtHours.Visible = False
        '
        'lblHours
        '
        Me.lblHours.Location = New System.Drawing.Point(560, 144)
        Me.lblHours.Name = "lblHours"
        Me.lblHours.Size = New System.Drawing.Size(48, 17)
        Me.lblHours.TabIndex = 255
        Me.lblHours.Text = "Hours:"
        Me.lblHours.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblHours.Visible = False
        '
        'cmbCourseTitle
        '
        Me.cmbCourseTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCourseTitle.DropDownWidth = 180
        Me.cmbCourseTitle.ItemHeight = 13
        Me.cmbCourseTitle.Location = New System.Drawing.Point(112, 40)
        Me.cmbCourseTitle.Name = "cmbCourseTitle"
        Me.cmbCourseTitle.Size = New System.Drawing.Size(296, 21)
        Me.cmbCourseTitle.TabIndex = 0
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(456, 312)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 23)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Close"
        '
        'ugCourseDates
        '
        Me.ugCourseDates.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCourseDates.Location = New System.Drawing.Point(64, 152)
        Me.ugCourseDates.Name = "ugCourseDates"
        Me.ugCourseDates.Size = New System.Drawing.Size(552, 152)
        Me.ugCourseDates.TabIndex = 256
        '
        'Courses
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(632, 349)
        Me.Controls.Add(Me.ugCourseDates)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.cmbCourseTitle)
        Me.Controls.Add(Me.txtHours)
        Me.Controls.Add(Me.txtCourseTitle)
        Me.Controls.Add(Me.txtCourseDates)
        Me.Controls.Add(Me.txtLocation)
        Me.Controls.Add(Me.chkActive)
        Me.Controls.Add(Me.lblHours)
        Me.Controls.Add(Me.cmbProvider)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.lblProvider)
        Me.Controls.Add(Me.lblCourseTitle)
        Me.Controls.Add(Me.lblCourseDates)
        Me.Controls.Add(Me.lblLocation)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnNew)
        Me.Name = "Courses"
        Me.Text = "Courses"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.ugCourseDates, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub LoadDisplayData()
        Try
            cmbProvider.DisplayMember = "Provider_Name"
            cmbProvider.ValueMember = "Provider_ID"
            cmbProvider.DataSource = oProvider.ListProviderNames(True)
            cmbProvider.Text = String.Empty


            'cmbType.DisplayMember = "Property_Name"
            'cmbType.ValueMember = "Property_ID"
            'cmbType.DataSource = oCourse.ListCourseTypes()
            'cmbType.SelectedIndex = -1
            'cmbType.SelectedIndex = -1
            'cmbType.Text = String.Empty

            cmbCourseTitle.DisplayMember = "Course_Title"
            cmbCourseTitle.ValueMember = "Course_ID"
            cmbCourseTitle.DataSource = oCourse.ListCourseTitles
            'cmbCourseTitle.SelectedIndex = -1
            'cmbCourseTitle.SelectedIndex = -1
            cmbCourseTitle.Text = String.Empty

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Function ValidateCoursedates() As Boolean
        Dim drRow As DataRow
        Dim dtColumn As DataColumn

        Try
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True
            If dtCourses.Rows.Count > 0 Then

                For Each drRow In dtCourses.Rows

                    For Each dtColumn In dtCourses.Columns
                        If drRow(dtColumn) Is DBNull.Value Then
                            errStr += "Provide " + dtColumn.ColumnName.ToString + " on row" + drRow("COURSEDATES_NUMBER").ToString + vbCrLf
                            validateSuccess = False
                        End If

                    Next

                Next
            End If


            'End If
            If errStr.Length > 0 Or Not validateSuccess Then
                MsgBox(errStr)
            End If
            Return validateSuccess
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub ClearCourseDatesGrid()
        Try
            dtCourses = New DataTable
            ugCourseDates.DataSource = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetupCourseDatesTable()
        Try
            dtCourses = oCourse.CourseDates.Clone
            dtCourses = oCourse.CourseDates
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub FillCourseDates()
        '  Dim i As Integer = 0
        Try
            'dtAnalysis = pClosure.SamplesTable
            ClearCourseDatesGrid()
            SetupCourseDatesTable()
            ugCourseDates.DataSource = Nothing

            dtCourses.DefaultView.Sort = "COURSEDATES_NUMBER"
            ugCourseDates.DataSource = dtCourses
            setUpCourseDates()
            nID = -1
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub setUpCourseDates()
        Try
            ugCourseDates.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True
            ugCourseDates.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom


            ugCourseDates.DisplayLayout.Bands(0).Columns("Course Dates").Width = 200
            ugCourseDates.DisplayLayout.Bands(0).Columns("Location").Width = 100
            ugCourseDates.DisplayLayout.Bands(0).Columns("Course Type").Width = 100
            ugCourseDates.DisplayLayout.Bands(0).Columns("Hours").Width = 50

            ugCourseDates.DisplayLayout.Bands(0).Columns("COURSE_ID").Hidden = True
            ugCourseDates.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
            ugCourseDates.DisplayLayout.Bands(0).Columns("COURSEDATES_ID").Hidden = True

            ugCourseDates.DisplayLayout.Bands(0).Columns("COURSEDATES_NUMBER").Hidden = True
            'ugCourseDates.DisplayLayout.Bands(0).Columns("DATE_CREATED").Hidden = True
            'ugCourseDates.DisplayLayout.Bands(0).Columns("LAST_EDITED_BY").Hidden = True
            'ugCourseDates.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True

            ugCourseDates.DisplayLayout.Bands(0).Columns("Course Type").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            If ugCourseDates.DisplayLayout.ValueLists.All.Length = 0 Then
                ugCourseDates.DisplayLayout.ValueLists.Add("CourseType")
            End If
            If ugCourseDates.DisplayLayout.Bands(0).Columns("Course Type").ValueList Is Nothing Then
                Dim vListCourseType As New Infragistics.Win.ValueList
                For Each dr As DataRow In oCourse.ListCourseTypes(True).Rows
                    vListCourseType.ValueListItems.Add(dr.Item("PROPERTY_ID"), dr.Item("PROPERTY_NAME").ToString)
                Next
                ugCourseDates.DisplayLayout.Bands(0).Columns("Course Type").ValueList = vListCourseType
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Courses_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True
            LoadDisplayData()
            bolLoading = False
            ClearCourseData()
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            bolLoading = True
            nCourseID = 0
            ClearCourseData()

            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbCourseTitle_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCourseTitle.SelectedIndexChanged
        Dim strOldProviderName As String = oProvider.ID
        If bolLoading Then Exit Sub

        Try
            If cmbCourseTitle.SelectedIndex = -1 Then Exit Sub

            oCourse.Retrieve(cmbCourseTitle.SelectedValue)
            'Course.Retrieve(cmbCourseTitle.SelectedValue.ToString)


            Me.txtCourseTitle.Text = cmbCourseTitle.Text
            Me.txtCourseTitle.Text = cmbCourseTitle.Text
            UIUtilsGen.SetComboboxItemByText(cmbProvider, oCourse.ProviderName)
            'Me.cmbProvider.SelectedValue = oCourse.ProviderID
            Me.chkActive.Checked = oCourse.Active
            'Me.txtCourseID.Text = cmbCourseTitle.SelectedValue
            nCourseID = cmbCourseTitle.SelectedValue
            'Me.txtHours.Text = oCourse.CreditHours
            'Me.txtLocation.Text = oCourse.Location
            'Me.txtCourseDates.Text = oCourse.CourseDates
            'Me.cmbType.SelectedValue = oCourse.CourseTypeID

            Me.FillCourseDates()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim drNew As DataRow
        Dim bolSuccess As Boolean
        Try
            'For Each drNew In dtCourses.Rows
            '    oCourse.GetCurrentInfo(drNew("Course_ID"))
            '    oCourse.Active = chkActive.Checked
            '    oCourse.CourseTitle = txtCourseTitle.Text
            'Next

            oCourseInfo = New MUSTER.Info.CourseInfo(nCourseID, _
                                   chkActive.Checked, _
                                   cmbProvider.SelectedValue, _
                                   txtCourseTitle.Text, _
                                   cmbProvider.Text, _
                                     False, _
                                   IIf(nCourseID <= 0, MusterContainer.AppUser.ID, ""), _
                                   Now, _
                                   IIf(nCourseID > 0, MusterContainer.AppUser.ID, ""), _
                                   CDate("01/01/0001"))

            oCourse.Add(oCourseInfo)
            If oCourse.Add(oCourseInfo) Then

                If oCourse.ID <= 0 Then
                    oCourse.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oCourse.ModifiedBy = MusterContainer.AppUser.ID
                End If

                oCourse.Save(CType(UIUtilsGen.ModuleID.CompanyAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If


                'oCourse.f()
                nCourseID = oCourse.ID
                If nCourseID > 0 Then

                    For Each drNew In dtCourses.Rows
                        drNew.Item("COURSE_ID") = nCourseID
                    Next
                    'If ValidateCoursedates() Then
                    bolSuccess = oCourse.SaveCourseDates(MusterContainer.AppUser.ID)
                    'Else
                    If Not bolSuccess Then
                        ugCourseDates.DisplayLayout.Bands(0).Columns("COURSEDATES_NUMBER").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        Exit Try
                    End If
                    If bolSuccess Then
                        MsgBox("Course Saved Successfully")

                    End If

                    bolLoading = True
                    cmbCourseTitle.DisplayMember = "Course_Title"
                    cmbCourseTitle.ValueMember = "Course_ID"
                    cmbCourseTitle.DataSource = oCourse.ListCourseTitles
                    bolLoading = False
                    If cmbCourseTitle.SelectedValue = oCourse.ID Then
                        FillCourseDates()
                    Else
                        'cmbCourseTitle.SelectedValue = oCourse.ID
                        UIUtilsGen.SetComboboxItemByValue(cmbCourseTitle, oCourse.ID)
                    End If

                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()

            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim drRow As DataRow
        Try
            If cmbCourseTitle.SelectedIndex = -1 Then
                MsgBox("Please select the course title to be deleted")
                Exit Sub
            End If
            Dim msgResult As MsgBoxResult = MsgBox("Are you sure you wish to DELETE the Course : " & cmbCourseTitle.Text & " ?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "DELETE COURSE")
            If msgResult = MsgBoxResult.Yes Then
                'Delete Provider 
                oCourse.Deleted = True
                If oCourse.ID <= 0 Then
                    oCourse.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oCourse.ModifiedBy = MusterContainer.AppUser.ID
                End If
                oCourse.Save(CType(UIUtilsGen.ModuleID.CompanyAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                oCourse.Remove(cmbCourseTitle.SelectedValue)
                oCourse.SaveCourseDates(MusterContainer.AppUser.ID)
                For Each drRow In dtCourses.Rows
                    drRow.Item("DELETED") = oCourse.Deleted
                Next
                oCourse.SaveCourseDates(MusterContainer.AppUser.ID)
                MsgBox("Course Deleted Successfully.")
                bolLoading = True
                cmbCourseTitle.DisplayMember = "Course_Title"
                cmbCourseTitle.ValueMember = "Course_ID"
                cmbCourseTitle.DataSource = oCourse.ListCourseTitles
                cmbCourseTitle.SelectedIndex = -1
                cmbCourseTitle.Text = String.Empty
                ClearCourseData()
                Me.cmbCourseTitle.Focus()
                bolLoading = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ClearCourseData()

        bolLoading = True

        'Me.txtCourseID.Text = String.Empty
        'Me.cmbProvider.SelectedIndex = -1
        Me.cmbCourseTitle.SelectedIndex = -1
        Me.cmbProvider.Text = String.Empty
        Me.txtCourseTitle.Text = String.Empty
        Me.txtCourseDates.Text = String.Empty
        Me.txtLocation.Text = String.Empty
        Me.cmbCourseTitle.SelectedIndex = -1
        Me.cmbType.SelectedIndex = -1
        Me.cmbType.Text = String.Empty
        Me.txtHours.Text = String.Empty
        Me.chkActive.Checked = True
        Me.cmbProvider.SelectedIndex = -1
        Me.cmbProvider.SelectedIndex = -1
        oCourse.CourseDates.Rows.Clear()
        FillCourseDates()
        bolLoading = False

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub oCourse_evtCourseErr(ByVal MsgStr As String) Handles oCourse.evtCourseErr
        If MsgStr <> String.Empty And MsgBox(MsgStr) = MsgBoxResult.OK Then
            MsgStr = String.Empty
            bolValidationFlg = True
            Exit Sub
        End If
    End Sub

    Private Sub ugCourseDates_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugCourseDates.AfterRowUpdate

        Try

            If e.Row.Update = True Then

                btnSave.Enabled = True
                oCourse.IsDirty = True

            End If
            If Not (e.Row.Cells("COURSE DATES").Text = String.Empty And e.Row.Cells("LOCATION").Text = String.Empty And e.Row.Cells("COURSE TYPE").Text = String.Empty And e.Row.Cells("HOURS").Text = String.Empty) Then
                If IsDBNull(dtCourses.Compute("Max([COURSEDATES_NUMBER])", "")) Then
                    e.Row.Cells("COURSEDATES_NUMBER").Value = 1
                Else
                    If IsDBNull(e.Row.Cells("COURSEDATES_NUMBER").Value) Then
                        e.Row.Cells("COURSEDATES_NUMBER").Value = dtCourses.Compute("Max([COURSEDATES_NUMBER])", "") + 1
                    End If
                End If
                If e.Row.Cells("Coursedates_id").Text = String.Empty Or e.Row.Cells("Coursedates_id").Text Is DBNull.Value Then
                    e.Row.Cells("Coursedates_id").Value = nID
                    nID -= 1
                    e.Row.Cells("COURSE_ID").Value = oCourse.ID
                    e.Row.Cells("DELETED").Value = False
                End If
            Else
                e.Row.Cells("COURSEDATES_NUMBER").Value = System.DBNull.Value
                e.Row.Cells("Coursedates_id").Value = System.DBNull.Value
                e.Row.Cells("DELETED").Value = System.DBNull.Value
            End If
            ugCourseDates.DisplayLayout.Bands(0).Columns("COURSEDATES_NUMBER").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
        'ugrow = ugCourseDates.ActiveRow
        'If IsDBNull(ugrow.Cells("COURSE_ID").Value) Then
        '    oCourse.GetCurrentInfo(0)
        '    ugrow.Cells("COURSE_ID").Value = oCourse.ID
        'Else
        '    oCourse.GetCurrentInfo(ugrow.Cells("COURSE_ID").Value)
        'End If

        'oCourse.ProviderID = cmbProvider.SelectedValue

        'If ugrow.Cells("Course Dates").Text <> String.Empty Then
        '    oCourse.CourseDates = ugrow.Cells("Course Dates").Value
        'Else
        '    oCourse.CourseDates = String.Empty
        'End If
        'If ugrow.Cells("Location").Text <> String.Empty Then
        '    oCourse.Location = ugrow.Cells("Location").Value
        'Else
        '    oCourse.Location = String.Empty
        'End If
        'If ugrow.Cells("Course Type").Text <> String.Empty Then
        '    oCourse.CourseTypeID = ugrow.Cells("Course Type").Value
        'Else
        '    oCourse.CourseTypeID = 0
        'End If
        'If ugrow.Cells("Hours").Text <> String.Empty Then
        '    oCourse.CreditHours = ugrow.Cells("Hours").Value
        'Else
        '    oCourse.CreditHours = 0
        'End If
        'oCourse.ProviderName = cmbProvider.Text
        'oCourse.CourseTitle = txtCourseTitle.Text
        'oCourse.Deleted = False


    End Sub

    Private Sub ugCourseDates_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles ugCourseDates.BeforeRowsDeleted
        Dim drow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            e.DisplayPromptMsg = False
            Dim results As MsgBoxResult = MsgBox("You have selected " + e.Rows.Length.ToString + "row(s) for deletion." + _
                                                    "Do you want to continue", MsgBoxStyle.YesNo, "Delete Row(s)")
            If results = MsgBoxResult.Yes Then
                For Each drow In e.Rows
                    drow.Cells("DELETED").Value = True
                    e.Cancel = True
                    ugCourseDates.ActiveRow = drow
                    ugCourseDates.ActiveRow.Hidden = True
                Next
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    
End Class
