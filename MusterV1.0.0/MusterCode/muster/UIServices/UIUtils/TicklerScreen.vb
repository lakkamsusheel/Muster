Public Class TickerScreen
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()


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
    Friend WithEvents lblTickerlist As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents pnlEditActivityPlan As System.Windows.Forms.Panel
    Friend WithEvents btncancelEdit As System.Windows.Forms.Button
    Friend WithEvents LblDescFrom As System.Windows.Forms.Label
    Friend WithEvents lblDescFor As System.Windows.Forms.Label
    Friend WithEvents LblDescSubject As System.Windows.Forms.Label
    Friend WithEvents LblFrom As System.Windows.Forms.Label
    Friend WithEvents LblFor As System.Windows.Forms.Label
    Friend WithEvents lblSubject As System.Windows.Forms.Label
    Friend WithEvents lblMsgDesc As System.Windows.Forms.Label
    Friend WithEvents txtMsg As System.Windows.Forms.TextBox
    Friend WithEvents BtnOpenEntity As System.Windows.Forms.Button
    Friend WithEvents LblEntityType As System.Windows.Forms.Label
    Friend WithEvents btnAttachment As System.Windows.Forms.Button
    Friend WithEvents ugticklerMessages As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnCompose As System.Windows.Forms.Button
    Friend WithEvents btnCompleted As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents btnReschedule_Apply As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ugticklerMessages = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lblTickerlist = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.pnlEditActivityPlan = New System.Windows.Forms.Panel
        Me.btnCompleted = New System.Windows.Forms.Button
        Me.btnAttachment = New System.Windows.Forms.Button
        Me.LblEntityType = New System.Windows.Forms.Label
        Me.BtnOpenEntity = New System.Windows.Forms.Button
        Me.txtMsg = New System.Windows.Forms.TextBox
        Me.lblMsgDesc = New System.Windows.Forms.Label
        Me.lblSubject = New System.Windows.Forms.Label
        Me.LblFor = New System.Windows.Forms.Label
        Me.LblFrom = New System.Windows.Forms.Label
        Me.LblDescSubject = New System.Windows.Forms.Label
        Me.lblDescFor = New System.Windows.Forms.Label
        Me.LblDescFrom = New System.Windows.Forms.Label
        Me.btncancelEdit = New System.Windows.Forms.Button
        Me.btnReschedule_Apply = New System.Windows.Forms.Button
        Me.btnCompose = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.TabPage3 = New System.Windows.Forms.TabPage
        CType(Me.ugticklerMessages, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlEditActivityPlan.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ugticklerMessages
        '
        Me.ugticklerMessages.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ugticklerMessages.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugticklerMessages.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugticklerMessages.DisplayLayout.Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
        Me.ugticklerMessages.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.None
        Me.ugticklerMessages.DisplayLayout.Override.AllowColSwapping = Infragistics.Win.UltraWinGrid.AllowColSwapping.NotAllowed
        Me.ugticklerMessages.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugticklerMessages.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.False
        Me.ugticklerMessages.DisplayLayout.Override.AllowGroupMoving = Infragistics.Win.UltraWinGrid.AllowGroupMoving.NotAllowed
        Me.ugticklerMessages.DisplayLayout.Override.AllowGroupSwapping = Infragistics.Win.UltraWinGrid.AllowGroupSwapping.NotAllowed
        Me.ugticklerMessages.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.False
        Me.ugticklerMessages.DisplayLayout.Override.AllowRowLayoutCellSizing = Infragistics.Win.UltraWinGrid.RowLayoutSizing.None
        Me.ugticklerMessages.DisplayLayout.Override.AllowRowLayoutLabelSizing = Infragistics.Win.UltraWinGrid.RowLayoutSizing.None
        Me.ugticklerMessages.DisplayLayout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.False
        Me.ugticklerMessages.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugticklerMessages.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugticklerMessages.Location = New System.Drawing.Point(16, 48)
        Me.ugticklerMessages.Name = "ugticklerMessages"
        Me.ugticklerMessages.Size = New System.Drawing.Size(920, 200)
        Me.ugticklerMessages.TabIndex = 47
        Me.ugticklerMessages.Text = "My Messages"
        '
        'lblTickerlist
        '
        Me.lblTickerlist.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTickerlist.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTickerlist.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTickerlist.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTickerlist.Location = New System.Drawing.Point(16, 32)
        Me.lblTickerlist.Name = "lblTickerlist"
        Me.lblTickerlist.Size = New System.Drawing.Size(920, 16)
        Me.lblTickerlist.TabIndex = 46
        Me.lblTickerlist.Text = "New Messages"
        Me.lblTickerlist.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(864, 472)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 45
        Me.btnClose.Text = "Close"
        '
        'pnlEditActivityPlan
        '
        Me.pnlEditActivityPlan.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlEditActivityPlan.AutoScroll = True
        Me.pnlEditActivityPlan.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlEditActivityPlan.Controls.Add(Me.btnCompleted)
        Me.pnlEditActivityPlan.Controls.Add(Me.btnAttachment)
        Me.pnlEditActivityPlan.Controls.Add(Me.LblEntityType)
        Me.pnlEditActivityPlan.Controls.Add(Me.BtnOpenEntity)
        Me.pnlEditActivityPlan.Controls.Add(Me.txtMsg)
        Me.pnlEditActivityPlan.Controls.Add(Me.lblMsgDesc)
        Me.pnlEditActivityPlan.Controls.Add(Me.lblSubject)
        Me.pnlEditActivityPlan.Controls.Add(Me.LblFor)
        Me.pnlEditActivityPlan.Controls.Add(Me.LblFrom)
        Me.pnlEditActivityPlan.Controls.Add(Me.LblDescSubject)
        Me.pnlEditActivityPlan.Controls.Add(Me.lblDescFor)
        Me.pnlEditActivityPlan.Controls.Add(Me.LblDescFrom)
        Me.pnlEditActivityPlan.Controls.Add(Me.btncancelEdit)
        Me.pnlEditActivityPlan.Controls.Add(Me.btnReschedule_Apply)
        Me.pnlEditActivityPlan.Enabled = False
        Me.pnlEditActivityPlan.Location = New System.Drawing.Point(16, 256)
        Me.pnlEditActivityPlan.Name = "pnlEditActivityPlan"
        Me.pnlEditActivityPlan.Size = New System.Drawing.Size(920, 208)
        Me.pnlEditActivityPlan.TabIndex = 44
        '
        'btnCompleted
        '
        Me.btnCompleted.Enabled = False
        Me.btnCompleted.Location = New System.Drawing.Point(8, 144)
        Me.btnCompleted.Name = "btnCompleted"
        Me.btnCompleted.Size = New System.Drawing.Size(128, 23)
        Me.btnCompleted.TabIndex = 57
        Me.btnCompleted.Text = "Set as Completed"
        '
        'btnAttachment
        '
        Me.btnAttachment.Enabled = False
        Me.btnAttachment.Location = New System.Drawing.Point(328, 0)
        Me.btnAttachment.Name = "btnAttachment"
        Me.btnAttachment.Size = New System.Drawing.Size(272, 23)
        Me.btnAttachment.TabIndex = 56
        Me.btnAttachment.Text = "View Attachment"
        '
        'LblEntityType
        '
        Me.LblEntityType.BackColor = System.Drawing.SystemColors.Desktop
        Me.LblEntityType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblEntityType.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.LblEntityType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblEntityType.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.LblEntityType.Location = New System.Drawing.Point(1, 176)
        Me.LblEntityType.Name = "LblEntityType"
        Me.LblEntityType.Size = New System.Drawing.Size(913, 24)
        Me.LblEntityType.TabIndex = 55
        Me.LblEntityType.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BtnOpenEntity
        '
        Me.BtnOpenEntity.Enabled = False
        Me.BtnOpenEntity.Location = New System.Drawing.Point(776, 144)
        Me.BtnOpenEntity.Name = "BtnOpenEntity"
        Me.BtnOpenEntity.Size = New System.Drawing.Size(128, 23)
        Me.BtnOpenEntity.TabIndex = 54
        Me.BtnOpenEntity.Text = "Open Entity"
        '
        'txtMsg
        '
        Me.txtMsg.BackColor = System.Drawing.SystemColors.HighlightText
        Me.txtMsg.Location = New System.Drawing.Point(80, 56)
        Me.txtMsg.Multiline = True
        Me.txtMsg.Name = "txtMsg"
        Me.txtMsg.ReadOnly = True
        Me.txtMsg.Size = New System.Drawing.Size(792, 80)
        Me.txtMsg.TabIndex = 53
        Me.txtMsg.Text = ""
        '
        'lblMsgDesc
        '
        Me.lblMsgDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMsgDesc.Location = New System.Drawing.Point(8, 56)
        Me.lblMsgDesc.Name = "lblMsgDesc"
        Me.lblMsgDesc.Size = New System.Drawing.Size(64, 16)
        Me.lblMsgDesc.TabIndex = 52
        Me.lblMsgDesc.Text = "Message :"
        Me.lblMsgDesc.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblSubject
        '
        Me.lblSubject.BackColor = System.Drawing.SystemColors.ControlLight
        Me.lblSubject.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSubject.Location = New System.Drawing.Point(80, 32)
        Me.lblSubject.Name = "lblSubject"
        Me.lblSubject.Size = New System.Drawing.Size(792, 16)
        Me.lblSubject.TabIndex = 51
        Me.lblSubject.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LblFor
        '
        Me.LblFor.BackColor = System.Drawing.SystemColors.ControlLight
        Me.LblFor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblFor.Location = New System.Drawing.Point(720, 8)
        Me.LblFor.Name = "LblFor"
        Me.LblFor.Size = New System.Drawing.Size(152, 16)
        Me.LblFor.TabIndex = 50
        Me.LblFor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LblFrom
        '
        Me.LblFrom.BackColor = System.Drawing.SystemColors.ControlLight
        Me.LblFrom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LblFrom.Location = New System.Drawing.Point(80, 8)
        Me.LblFrom.Name = "LblFrom"
        Me.LblFrom.Size = New System.Drawing.Size(144, 16)
        Me.LblFrom.TabIndex = 49
        Me.LblFrom.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LblDescSubject
        '
        Me.LblDescSubject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblDescSubject.Location = New System.Drawing.Point(8, 32)
        Me.LblDescSubject.Name = "LblDescSubject"
        Me.LblDescSubject.Size = New System.Drawing.Size(64, 16)
        Me.LblDescSubject.TabIndex = 48
        Me.LblDescSubject.Text = "Subject :"
        Me.LblDescSubject.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblDescFor
        '
        Me.lblDescFor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescFor.Location = New System.Drawing.Point(664, 8)
        Me.lblDescFor.Name = "lblDescFor"
        Me.lblDescFor.Size = New System.Drawing.Size(48, 16)
        Me.lblDescFor.TabIndex = 47
        Me.lblDescFor.Text = "For :"
        Me.lblDescFor.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'LblDescFrom
        '
        Me.LblDescFrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblDescFrom.Location = New System.Drawing.Point(8, 8)
        Me.LblDescFrom.Name = "LblDescFrom"
        Me.LblDescFrom.Size = New System.Drawing.Size(64, 16)
        Me.LblDescFrom.TabIndex = 46
        Me.LblDescFrom.Text = "From :"
        Me.LblDescFrom.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btncancelEdit
        '
        Me.btncancelEdit.BackColor = System.Drawing.Color.Brown
        Me.btncancelEdit.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.btncancelEdit.Location = New System.Drawing.Point(888, 0)
        Me.btncancelEdit.Name = "btncancelEdit"
        Me.btncancelEdit.Size = New System.Drawing.Size(24, 23)
        Me.btncancelEdit.TabIndex = 45
        Me.btncancelEdit.Text = "X"
        '
        'btnReschedule_Apply
        '
        Me.btnReschedule_Apply.Location = New System.Drawing.Point(144, 144)
        Me.btnReschedule_Apply.Name = "btnReschedule_Apply"
        Me.btnReschedule_Apply.Size = New System.Drawing.Size(168, 23)
        Me.btnReschedule_Apply.TabIndex = 57
        Me.btnReschedule_Apply.Text = "Re-schedule Posted Message"
        '
        'btnCompose
        '
        Me.btnCompose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCompose.Location = New System.Drawing.Point(16, 472)
        Me.btnCompose.Name = "btnCompose"
        Me.btnCompose.Size = New System.Drawing.Size(128, 23)
        Me.btnCompose.TabIndex = 55
        Me.btnCompose.Text = "Compose Message"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(16, 8)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(156, 24)
        Me.TabControl1.TabIndex = 56
        '
        'TabPage1
        '
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(148, 0)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "InBox"
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(148, -2)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Completed "
        '
        'TabPage3
        '
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(148, -2)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Outbox"
        '
        'TickerScreen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(948, 517)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnCompose)
        Me.Controls.Add(Me.ugticklerMessages)
        Me.Controls.Add(Me.lblTickerlist)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.pnlEditActivityPlan)
        Me.MaximumSize = New System.Drawing.Size(956, 574)
        Me.MinimumSize = New System.Drawing.Size(956, 320)
        Me.Name = "TickerScreen"
        Me.Text = "My Tickler Messages"
        CType(Me.ugticklerMessages, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlEditActivityPlan.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "private members"

    Private WithEvents _container As MusterContainer
    Private WithEvents _tm As TicklerManager
    Private _lastID As String
    Private _showCompleted As Boolean = False
    Private _showSent As Boolean = False

#End Region


#Region "Public properties"

    Public Property tm() As TicklerManager
        Get
            If _tm Is Nothing Then
                _tm = New TicklerManager
            End If
            Return _tm
        End Get
        Set(ByVal Value As TicklerManager)
            _tm = Value
        End Set
    End Property

    Public ReadOnly Property UnReadCount() As Integer

        Get
            If Not tm Is Nothing Then
                Return tm.UnReadMessages
            End If

            Return 0
        End Get

    End Property

#End Region


#Region "Public Events"

    Public InvokeRefreshScreen As New MethodInvoker(AddressOf CheckForUpdatesManual)
    Public Event OpenObject(ByVal moduleID As String, ByVal objectID As String, ByVal keyword As String)
    Public Event DisplayAlert(ByVal forced As Boolean)


#End Region

#Region "construct"

    Sub New(ByVal muster As MusterContainer)

        Me.New()

        _container = muster

    End Sub


#End Region

#Region "private methods"


    Private Sub populateMsgData(Optional ByVal completed As Boolean = False)



        If Not (ugticklerMessages.ActiveRow Is Nothing) Then
            With ugticklerMessages.ActiveRow

                tm.Message.Retrieve(.Cells("MsgID").Value, Not TabControl1.SelectedTab Is TabPage3, completed)

                If IsNumeric(tm.Message.ID) AndAlso Not tm.Message.Completed Then
                    Me.btnCompleted.Enabled = True
                Else
                    Me.btnCompleted.Enabled = False
                End If

                LblFor.Text = .Cells("ToID").Value
                LblFrom.Text = .Cells("FromID").Value

                With tm.Message

                    txtMsg.Text = .Message
                    lblSubject.Text = .Subject

                    If Not .ImageFile Is Nothing AndAlso .ImageFile.Length > 4 Then
                        btnAttachment.Enabled = True
                    Else
                        btnAttachment.Enabled = False
                    End If


                    If Not .ObjectID Is Nothing AndAlso .ObjectID.Length > 0 Then

                        BtnOpenEntity.Enabled = True
                        If .Keyword.IndexOf("ReportView") <> -1 Then

                            LblEntityType.Text = String.Format("Click Open Entity to open report '{0}'", .ObjectID)


                        ElseIf .Keyword.IndexOf("DocumentManager") <> -1 Then

                            If .ObjectID.IndexOf(";") > 0 Then

                                Dim user As String = .ObjectID
                                Dim year As String = "2009"

                                year = user.Substring(user.IndexOf(";") + 1)
                                user = user.Substring(0, user.IndexOf(";"))

                                LblEntityType.Text = String.Format("Click Open Entity to open the document manager under user {0} for year {1} ", user, year)

                            Else
                                LblEntityType.Text = String.Format("Click Open Entity to open the document manager")
                            End If



                        Else

                            Me.LblEntityType.Text = String.Format("Click Open Entity to open {0} '{1}' within the {2} module", .Keyword, _
                                                                  .ObjectID, ugticklerMessages.ActiveRow.Cells("ModuleDesc").Value)

                        End If
                    Else

                        BtnOpenEntity.Enabled = False
                        Me.LblEntityType.Text = String.Empty
                    End If

                    pnlEditActivityPlan.Enabled = True

                    If TabControl1.SelectedTab Is TabPage3 Then

                        If .PostDate > New Date(1910, 1, 1) Then
                            btnReschedule_Apply.Enabled = True
                        Else
                            btnReschedule_Apply.Enabled = False
                        End If

                        btnReschedule_Apply.Text = "Reschedule Posted Message"
                    Else

                        If .FromID <> "SYSTEM" And IsNumeric(.ID) Then
                            btnReschedule_Apply.Enabled = True
                        Else
                            btnReschedule_Apply.Enabled = False
                        End If

                        btnReschedule_Apply.Text = "Reply to Message"
                    End If

                End With

            End With
        End If


    End Sub

    Public Sub LoadticklerList(ByVal ds As DataTable)


        ugticklerMessages.DataSource = ds.DefaultView

        If Not ds Is Nothing AndAlso ds.Rows.Count > 0 Then

            For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In Me.ugticklerMessages.Rows
                checkRow(row)
            Next

        End If


        ugticklerMessages.Refresh()

    End Sub

    Sub tabSet()

        _showSent = False
        _showCompleted = False
        btnCompleted.Visible = True

        lblTickerlist.Text = "New Messages"



        If TabControl1.SelectedTab Is TabPage2 Then
            _showCompleted = True
            lblTickerlist.Text = "Completed & Read Messages"
            btnCompleted.Visible = False
        ElseIf TabControl1.SelectedTab Is TabPage3 Then
            _showSent = True
            lblTickerlist.Text = "Sent Messages"
            btnCompleted.Visible = False

        End If

        pnlEditActivityPlan.Enabled = False

    End Sub


#End Region

#Region "public Form Events"

    Private Sub pnlEditActivityPlan_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlEditActivityPlan.EnabledChanged

        If Not pnlEditActivityPlan.Enabled Then
            txtMsg.Text = String.Empty
            LblFor.Text = String.Empty
            lblSubject.Text = String.Empty
            LblFrom.Text = String.Empty

            btnReschedule_Apply.Enabled = False
            btnCompleted.Enabled = False

        End If
    End Sub

    Private Sub CheckForUpdates(ByVal ID As String) Handles _container.StartTicklerScreen

        _lastID = ID
        tm.refreshMessages(ID, False, _showCompleted, _showSent)

    End Sub

    Private Sub CheckForUpdatesManual()

        tm.refreshMessages(_lastID, True, _showCompleted, _showSent)

    End Sub

    Private Sub ShowMe(ByVal ds As DataTable, ByVal forced As Boolean) Handles _tm.NewFound

        If Not ds Is Nothing Then

            _showCompleted = False
            _showSent = False

            LoadticklerList(ds)

            If Not Visible Then

                RaiseEvent DisplayAlert(forced)

            Else


                ugTicklerMessages_InitializeLayout(Me.ugticklerMessages, Nothing)

            End If
        End If


    End Sub

    Private Sub TryToOpenMessage(ByVal completed As Boolean)

        If Not ugticklerMessages.ActiveRow Is Nothing Then

            Try
                populateMsgData(completed)

            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            End Try


        Else
            pnlEditActivityPlan.Enabled = False

        End If

    End Sub

    Private Sub tabControlChange(ByVal sender As Object, ByVal e As EventArgs) Handles TabControl1.SelectedIndexChanged

        tabSet()

        If Not tm Is Nothing AndAlso Not MusterContainer.AppUser Is Nothing Then
            tm.refreshMessages(MusterContainer.AppUser.ID, True, _showCompleted, _showSent)
        End If


    End Sub
    Private Sub ugTicklerMessages_AfterRowActivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugticklerMessages.Click

        TryToOpenMessage(False)
    End Sub

    Private Sub btnCompleted_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCompleted.Click

        TryToOpenMessage(True)
        tm.refreshMessages(MusterContainer.AppUser.ID, True)

    End Sub




    Private Sub BtnOpenEntity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOpenEntity.Click

        If Not Me.ugticklerMessages.ActiveRow Is Nothing Then
            With Me.ugticklerMessages.ActiveRow
                _container.OpenEntityFromTickler(.Cells("ModuleDesc").Value.ToString, .Cells("ObjectID").Value.ToString, .Cells("keyword").Value.ToString)

                Hide()

            End With
        End If

    End Sub

    Private Sub btnCompose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompose.Click

        Dim newtickler As New NewTicklerMessageScreen

        newtickler.ShowDialog()
        newtickler.Dispose()
        newtickler = Nothing
    End Sub


    Private Sub btnReschedule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReschedule_Apply.Click

        Dim newtickler As New NewTicklerMessageScreen

        If TabControl1.SelectedTab Is TabPage3 Then
            newtickler.Tickler = tm.Message
        Else
            Dim newMessage As New BusinessLogic.pTicklerMessage
            newMessage.Subject = String.Format("RE: {0}", tm.Message.Subject)
            newMessage.ObjectID = tm.Message.ObjectID
            newMessage.Keyword = tm.Message.Keyword
            newMessage.ToID = tm.Message.FromID
            newMessage.FromID = _container.AppUser.ID
            newMessage.ModuleID = tm.Message.ModuleID
            newMessage.ID = "REPLY"
            newMessage.PostDate = Nothing
            newMessage.ImageFile = tm.Message.ImageFile
            newMessage.IsIssue = tm.Message.IsIssue

            newtickler.Tickler = newMessage

        End If



        newtickler.ShowDialog()

        newtickler.Dispose()
        newtickler = Nothing

        tabControlChange(TabControl1, Nothing)

    End Sub


    Public Sub ShowActualForm(ByVal ta As TicklerAlert)

        ta.Close()

        Me.ShowDialog()

    End Sub

    Public Sub BtnCloseClicked(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click

        pnlEditActivityPlan.Enabled = False
        Close()
    End Sub

    Private Sub btncancelEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncancelEdit.Click
        pnlEditActivityPlan.Enabled = False
    End Sub

    Private Sub btnAttachment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAttachment.Click

        With tm.Message

            If .ImageFile.Length > 4 AndAlso System.IO.File.Exists(.ImageFile) Then
                Dim picform As New Form
                Dim img As Image
                img = Image.FromFile(.ImageFile)


                picform.BackgroundImage = img
                picform.Size = New Size(800, 600)
                picform.Opacity = 0.8



                If Not Me.ugticklerMessages.ActiveRow Is Nothing Then
                    picform.Text = String.Format("Attached picture for MUSTER Tickler Message {0}", Me.ugticklerMessages.ActiveRow.Cells("RowNum").Value)
                Else
                    picform.Text = "Attched to picture to a MUSTER tickler message"
                End If

                picform.WindowState = FormWindowState.Normal
                picform.Show()

            Else

                MsgBox("The attachment file for this message no longer exists in the server's directories")

            End If

        End With

    End Sub

#End Region


#Region "Grid Setup"

    Private Sub checkRow(ByVal e As Infragistics.Win.UltraWinGrid.UltraGridRow)

        e.Appearance.FontData.SizeInPoints = 8

        If e.Cells("IsNew").Value = "true" AndAlso e.Cells("MsgRead").Value = "0" Then
            e.Appearance.ForeColor = Color.Blue
        End If

        If e.Cells("MsgRead").Value.ToString = "0" Then

            e.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
        Else
            e.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.False
        End If

    End Sub


    Private Sub ugTicklerMessages_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugticklerMessages.InitializeLayout


        Dim layout As Infragistics.Win.UltraWinGrid.UltraGridLayout

        If e Is Nothing Then
            layout = ugticklerMessages.DisplayLayout
        Else
            layout = e.Layout

        End If

        With layout.Bands(0)

            Try
                layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect



                .Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                .Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No

                .Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
                .Columns("MsgID").Hidden = True
                .Columns("Msg").Hidden = True
                .Columns("MsgRead").Hidden = True
                .Columns("Completed").Hidden = True
                .Columns("IsNew").Hidden = True
                .Columns("ModuleGroupID").Hidden = True
                .Columns("ModuleDesc").Hidden = True
                .Columns("ObjectID").Hidden = True
                .Columns("Keyword").Hidden = True

                .Columns("RowNum").Header.Caption = "#"
                .Columns("FromID").Header.Caption = "From"
                .Columns("ToID").Header.Caption = "To"
                .Columns("Subject").Header.Caption = "Subject"
                .Columns("DatePulled").Header.Caption = "Date of Message"




                layout.AutoFitColumns = True
                .Columns("RowNum").Width = 15
                .Columns("DatePulled").Width = 80

                .Columns("FromID").Width = 100
                .Columns("ToID").Width = 100
                .Columns("Subject").Width = 400


                .Columns("RowNum").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                .Columns("FromID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                .Columns("ToID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                .Columns("Subject").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit


                .Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
                .Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.Default

            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()

            End Try


        End With

    End Sub

#End Region






End Class
