Public Class NewTicklerMessageScreen
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
    Friend WithEvents lblDescFor As System.Windows.Forms.Label
    Friend WithEvents CmbTo As System.Windows.Forms.ComboBox
    Friend WithEvents LblDescSubj As System.Windows.Forms.Label
    Friend WithEvents TxtSubject As System.Windows.Forms.TextBox
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents LblDsecMsg As System.Windows.Forms.Label
    Friend WithEvents lblModuleComboDesc As System.Windows.Forms.Label
    Public WithEvents cmbSearchModule As System.Windows.Forms.ComboBox
    Friend WithEvents cmbQuickSearchFilter As System.Windows.Forms.ComboBox
    Friend WithEvents LblFilter As System.Windows.Forms.Label
    Friend WithEvents lbDescID As System.Windows.Forms.Label
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents lblPostDesc As System.Windows.Forms.Label
    Friend WithEvents dtPost As System.Windows.Forms.DateTimePicker
    Friend WithEvents cboxTabPage As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblDescFor = New System.Windows.Forms.Label
        Me.CmbTo = New System.Windows.Forms.ComboBox
        Me.LblDescSubj = New System.Windows.Forms.Label
        Me.TxtSubject = New System.Windows.Forms.TextBox
        Me.txtMessage = New System.Windows.Forms.TextBox
        Me.LblDsecMsg = New System.Windows.Forms.Label
        Me.LblFilter = New System.Windows.Forms.Label
        Me.lbDescID = New System.Windows.Forms.Label
        Me.lblModuleComboDesc = New System.Windows.Forms.Label
        Me.cmbSearchModule = New System.Windows.Forms.ComboBox
        Me.cmbQuickSearchFilter = New System.Windows.Forms.ComboBox
        Me.txtID = New System.Windows.Forms.TextBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.lblPostDesc = New System.Windows.Forms.Label
        Me.dtPost = New System.Windows.Forms.DateTimePicker
        Me.cboxTabPage = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'lblDescFor
        '
        Me.lblDescFor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescFor.Location = New System.Drawing.Point(32, 8)
        Me.lblDescFor.Name = "lblDescFor"
        Me.lblDescFor.Size = New System.Drawing.Size(32, 16)
        Me.lblDescFor.TabIndex = 51
        Me.lblDescFor.Text = "For :"
        Me.lblDescFor.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CmbTo
        '
        Me.CmbTo.Location = New System.Drawing.Point(72, 8)
        Me.CmbTo.Name = "CmbTo"
        Me.CmbTo.Size = New System.Drawing.Size(120, 21)
        Me.CmbTo.TabIndex = 0
        Me.CmbTo.Text = "ComboBox1"
        '
        'LblDescSubj
        '
        Me.LblDescSubj.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblDescSubj.Location = New System.Drawing.Point(232, 8)
        Me.LblDescSubj.Name = "LblDescSubj"
        Me.LblDescSubj.Size = New System.Drawing.Size(56, 16)
        Me.LblDescSubj.TabIndex = 52
        Me.LblDescSubj.Text = "Subject :"
        Me.LblDescSubj.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtSubject
        '
        Me.TxtSubject.Location = New System.Drawing.Point(296, 8)
        Me.TxtSubject.Name = "TxtSubject"
        Me.TxtSubject.Size = New System.Drawing.Size(672, 20)
        Me.TxtSubject.TabIndex = 1
        Me.TxtSubject.Text = ""
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(72, 80)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(896, 72)
        Me.txtMessage.TabIndex = 2
        Me.txtMessage.Text = ""
        '
        'LblDsecMsg
        '
        Me.LblDsecMsg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblDsecMsg.Location = New System.Drawing.Point(0, 80)
        Me.LblDsecMsg.Name = "LblDsecMsg"
        Me.LblDsecMsg.Size = New System.Drawing.Size(64, 16)
        Me.LblDsecMsg.TabIndex = 53
        Me.LblDsecMsg.Text = "Message :"
        Me.LblDsecMsg.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LblFilter
        '
        Me.LblFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblFilter.Location = New System.Drawing.Point(488, 160)
        Me.LblFilter.Name = "LblFilter"
        Me.LblFilter.Size = New System.Drawing.Size(72, 16)
        Me.LblFilter.TabIndex = 59
        Me.LblFilter.Text = "Entity Type :"
        Me.LblFilter.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lbDescID
        '
        Me.lbDescID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbDescID.Location = New System.Drawing.Point(8, 160)
        Me.lbDescID.Name = "lbDescID"
        Me.lbDescID.Size = New System.Drawing.Size(56, 16)
        Me.lbDescID.TabIndex = 58
        Me.lbDescID.Text = "Entity ID :"
        Me.lbDescID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblModuleComboDesc
        '
        Me.lblModuleComboDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModuleComboDesc.Location = New System.Drawing.Point(264, 160)
        Me.lblModuleComboDesc.Name = "lblModuleComboDesc"
        Me.lblModuleComboDesc.Size = New System.Drawing.Size(56, 16)
        Me.lblModuleComboDesc.TabIndex = 57
        Me.lblModuleComboDesc.Text = "Module :"
        Me.lblModuleComboDesc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbSearchModule
        '
        Me.cmbSearchModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSearchModule.Location = New System.Drawing.Point(328, 160)
        Me.cmbSearchModule.MaxDropDownItems = 9
        Me.cmbSearchModule.Name = "cmbSearchModule"
        Me.cmbSearchModule.Size = New System.Drawing.Size(120, 21)
        Me.cmbSearchModule.TabIndex = 55
        '
        'cmbQuickSearchFilter
        '
        Me.cmbQuickSearchFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbQuickSearchFilter.Location = New System.Drawing.Point(568, 160)
        Me.cmbQuickSearchFilter.Name = "cmbQuickSearchFilter"
        Me.cmbQuickSearchFilter.Size = New System.Drawing.Size(120, 21)
        Me.cmbQuickSearchFilter.TabIndex = 56
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(72, 160)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(144, 20)
        Me.txtID.TabIndex = 3
        Me.txtID.Text = ""
        Me.txtID.WordWrap = False
        '
        'CheckBox1
        '
        Me.CheckBox1.Enabled = False
        Me.CheckBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(808, 160)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(160, 24)
        Me.CheckBox1.TabIndex = 60
        Me.CheckBox1.Text = "Add Image from Clipboard"
        Me.CheckBox1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox1.Location = New System.Drawing.Point(704, 184)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(264, 152)
        Me.PictureBox1.TabIndex = 61
        Me.PictureBox1.TabStop = False
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(872, 360)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(96, 23)
        Me.btnClose.TabIndex = 62
        Me.btnClose.Text = "Cancel"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(704, 360)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(96, 23)
        Me.Button1.TabIndex = 63
        Me.Button1.Text = "Save"
        '
        'lblPostDesc
        '
        Me.lblPostDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPostDesc.Location = New System.Drawing.Point(8, 40)
        Me.lblPostDesc.Name = "lblPostDesc"
        Me.lblPostDesc.Size = New System.Drawing.Size(56, 16)
        Me.lblPostDesc.TabIndex = 64
        Me.lblPostDesc.Text = "post :"
        Me.lblPostDesc.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtPost
        '
        Me.dtPost.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPost.Location = New System.Drawing.Point(72, 40)
        Me.dtPost.Name = "dtPost"
        Me.dtPost.ShowCheckBox = True
        Me.dtPost.Size = New System.Drawing.Size(120, 20)
        Me.dtPost.TabIndex = 65
        '
        'cboxTabPage
        '
        Me.cboxTabPage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboxTabPage.Location = New System.Drawing.Point(288, 40)
        Me.cboxTabPage.Name = "cboxTabPage"
        Me.cboxTabPage.Size = New System.Drawing.Size(680, 24)
        Me.cboxTabPage.TabIndex = 66
        Me.cboxTabPage.Text = "Tab"
        Me.cboxTabPage.Visible = False
        '
        'NewTicklerMessageScreen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(976, 390)
        Me.Controls.Add(Me.cboxTabPage)
        Me.Controls.Add(Me.dtPost)
        Me.Controls.Add(Me.lblPostDesc)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.LblFilter)
        Me.Controls.Add(Me.lbDescID)
        Me.Controls.Add(Me.lblModuleComboDesc)
        Me.Controls.Add(Me.cmbSearchModule)
        Me.Controls.Add(Me.cmbQuickSearchFilter)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.LblDsecMsg)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.TxtSubject)
        Me.Controls.Add(Me.LblDescSubj)
        Me.Controls.Add(Me.CmbTo)
        Me.Controls.Add(Me.lblDescFor)
        Me.Name = "NewTicklerMessageScreen"
        Me.Text = "New Tickler Message Wizard"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "private members"

    Private _pTickler As BusinessLogic.pTicklerMessage
    Private _PostMessage As Boolean = False
    Private _image As Image
#End Region

#Region "Public members"

    Public Property Tickler() As BusinessLogic.pTicklerMessage

        Get

            If _pTickler Is Nothing Then
                _pTickler = New BusinessLogic.pTicklerMessage

                _pTickler.Reset()

                _pTickler.PostDate = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)

            End If

            Return _pTickler

        End Get

        Set(ByVal Value As BusinessLogic.pTicklerMessage)
            _pTickler = Value
        End Set
    End Property

#End Region



#Region "Tickler Methods"

    Private Sub SetAsPostUpdate()

        If Tickler.ID <> "REPLY" Then
            _PostMessage = True
        Else
            dtPost.Enabled = False
        End If

        txtID.Enabled = False
        cmbSearchModule.Enabled = False
        cmbQuickSearchFilter.Enabled = False
        CmbTo.Enabled = False
        CheckBox1.Enabled = False
        cboxTabPage.Enabled = False

    End Sub

    Function storePictureInServer() As Boolean

        If Not Me.CheckBox1.Enabled Then
            Return True
        End If

        Dim retVal As Boolean = True
        Dim strFilename As String = String.Format("Attach_{0}_(1)_{2}_{3}_{4}_{5}_{6}_{7}", Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Tickler.FromID, Tickler.ToID, Tickler.ObjectID)
        Dim strServerShare As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue
        Dim path As String = String.Empty

        path = String.Format("{0}{1}{2}.bmp", strServerShare, IO.Path.DirectorySeparatorChar, strFilename)


        Dim lData As IDataObject = Clipboard.GetDataObject()

        If Not lData Is Nothing AndAlso lData.GetDataPresent(DataFormats.Bitmap) Then
            Dim lPictureBox As New PictureBox
            lPictureBox.Image = lData.GetData(DataFormats.Bitmap, True)

            lPictureBox.Image.Save(path, System.Drawing.Imaging.ImageFormat.Bmp)

            Tickler.ImageFile = path

            lPictureBox.Dispose()
        Else
            MsgBox("Clipboard data is not in bitmap form. Please the screen/view in question as press print screen again")
            retVal = False
        End If

        lData = Nothing

        Return retVal


    End Function

    Private Sub BindTicklerData()


        Tickler.FromID = MusterContainer.AppUser.ID
        Tickler.ImageFile = String.Empty
        Tickler.IsIssue = False


        If Tickler.PostDate = Nothing OrElse Tickler.PostDate < New Date(1910, 1, 1) Then
            Tickler.PostDate = dtPost.MinDate
            dtPost.Checked = False
        End If


        If Tickler.ID = String.Empty Then

            Tickler.ModuleID = _container.cmbSearchModule.SelectedValue
            Tickler.Keyword = _container.cmbQuickSearchFilter.SelectedValue
            Tickler.ObjectID = _container.txtOwnerQSKeyword.Text


        End If


        txtMessage.DataBindings.Add("Text", Tickler, "Message")

        TxtSubject.DataBindings.Add("Text", Tickler, "Subject")
        CmbTo.DataBindings.Add("SelectedValue", Tickler, "ToID")
        cmbSearchModule.DataBindings.Add("SelectedValue", Tickler, "ModuleID")
        cmbQuickSearchFilter.DataBindings.Add("SelectedValue", Tickler, "Keyword")
        txtID.DataBindings.Add("Text", Tickler, "ObjectID")
        dtPost.DataBindings.Add("Value", Tickler, "PostDate")



        If Tickler.ID = String.Empty Then
            Me.txtMessage.Text = "Enter New Message"
            txtMessage.Focus()
            Me.CmbTo.Focus()
        End If


        Dim lData As IDataObject = Clipboard.GetDataObject()

        If Not lData Is Nothing AndAlso lData.GetDataPresent(DataFormats.Bitmap) Then
            CheckBox1.Enabled = True

            _image = lData.GetData(DataFormats.Bitmap, True)

        Else
            CheckBox1.Enabled = False
        End If

        lData = Nothing

        If dtPost.Value < New Date(1910, 1, 1) Then
            dtPost.Checked = False
            dtPost.Refresh()
        End If

    End Sub


#End Region

#Region "Form Events"

    Sub LoadForm(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        If Not _pTickler Is Nothing AndAlso _pTickler.ID <> String.Empty Then
            SetAsPostUpdate()
        End If


        LoadCmbSearchModule()
        LoadToBox()
        BindTicklerData()


        If Tickler.ID = "REPLY" Then
            Tickler.ID = String.Empty
        End If



    End Sub

    Private Sub dtPost_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPost.ValueChanged

        If dtPost.Value <= dtPost.MinDate Then
            Text = "New Tickler Message Wizard"
        Else
            Text = "Posted Tickler Message Wizard"
        End If

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click

        _pTickler = Nothing
        Close()

    End Sub

    Private Sub SetUpTicklerExtras()

        If cboxTabPage.Visible AndAlso cboxTabPage.Checked Then
            Tickler.ObjectID = String.Format("{0};{1}", Tickler.ObjectID, "tbPageOwnerCitations")
        End If

    End Sub

    Private Sub NewTicklerMessageScreen_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If Tickler.IsDirty AndAlso MsgBox("Incomplete Message: Are you sure you want to cancel this message? ", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            e.Cancel = True
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        With PictureBox1

            If CheckBox1.Checked Then

                .SizeMode = PictureBoxSizeMode.StretchImage
                .Image = _image
            Else
                .Image = Nothing
            End If

        End With



    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim retVal As String = String.Empty

        Try
            If Not Me._PostMessage Then
                If storePictureInServer() Then


                    SetUpTicklerExtras()

                    TicklerManager.saveData(Tickler)

                    MsgBox("Message saved and sent!")

                    Close()

                End If
            Else

                TicklerManager.saveData(Tickler)

                If retVal.Length > 0 Then
                    Throw New Exception(String.Format("Error Updating Posted Tickler Message: {0}", retVal))
                End If

                MsgBox("Posted Message updated!")

                Close()

            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try



    End Sub

#End Region

#Region "Combo Box Events/Methods"

    Private Sub LoadToBox()

        Dim dtTable As DataTable = MusterContainer.AppUser.ListUsersToSend()
        CmbTo.DataSource = dtTable
        CmbTo.ValueMember = "PropertyValue"
        CmbTo.DisplayMember = "PropertyName"

    End Sub

    Private Sub LoadCmbSearchModule()

        Dim dtTable As DataTable = MusterContainer.AppUser.ListModulesUserCanSearch(MusterContainer.AppUser.UserKey)
        dtTable.DefaultView.Sort = "PROPERTY_NAME"
        cmbSearchModule.DataSource = dtTable.DefaultView
        cmbSearchModule.ValueMember = "PROPERTY_ID"
        cmbSearchModule.DisplayMember = "PROPERTY_NAME"

    End Sub

    Private Sub cmbSearchModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSearchModule.SelectedIndexChanged

        Dim oQS As New BusinessLogic.pSearch

        Dim dtTable As DataTable
        Try
            If TypeOf cmbSearchModule.SelectedValue Is Integer Then
                dtTable = oQS.PopulateQuickSearchFilter(cmbSearchModule.SelectedValue.ToString)
                cmbQuickSearchFilter.DataSource = dtTable
                If Not dtTable Is Nothing Then
                    cmbQuickSearchFilter.ValueMember = "PROPERTY_NAME"
                    cmbQuickSearchFilter.DisplayMember = "PROPERTY_NAME"
                    cmbQuickSearchFilter.SelectedIndex = 1
                End If
            End If

            If TypeOf cmbSearchModule.SelectedValue Is Integer AndAlso cmbSearchModule.SelectedValue = 613 Then
                cboxTabPage.Visible = True
                cboxTabPage.Text = "Send User to OCE tab"

            Else
                Me.cboxTabPage.Visible = False
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            oQS = Nothing
        End Try
    End Sub
#End Region


    Private Sub dtPost_Validating(ByVal sender As Object, ByVal e As EventArgs) Handles dtPost.Validated
        If Not dtPost.Checked Then
            dtPost.Value = dtPost.MinDate
        End If


    End Sub
End Class
