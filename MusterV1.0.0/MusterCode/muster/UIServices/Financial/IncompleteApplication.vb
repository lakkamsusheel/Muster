Public Class IncompleteApplication
    Inherits System.Windows.Forms.Form

    'Fixes
    '1.1        Thomas Franey         2/20/2009              Fixed close button to close
    '                                                        line: 368
    '1.1        Thomas Franey         2/20/2009              Added Object exists check for calling from tag in save Procedure 
    '                                                        line: 337   
    '1.2        Thoams Franey         3/05/2009              Added Commitment NULL logic with checkbox, also fix view to get only open balance
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
    Friend WithEvents pnlIncompleteApplicationBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlIncompleteApplicationDetails As System.Windows.Forms.Panel
    Friend WithEvents lblReceived As System.Windows.Forms.Label
    Friend WithEvents dtPickReceived As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblRequested As System.Windows.Forms.Label
    Friend WithEvents txtRequested As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents chkIncompleteApp As System.Windows.Forms.CheckBox
    Friend WithEvents ReasonPanel As System.Windows.Forms.Panel
    Friend WithEvents clbIncompleteReason As System.Windows.Forms.CheckedListBox
    Friend WithEvents txtOther As System.Windows.Forms.TextBox
    Friend WithEvents lblOther As System.Windows.Forms.Label
    Friend WithEvents lblIncompleteApplication As System.Windows.Forms.Label
    Friend WithEvents LblCommitment As System.Windows.Forms.Label
    Friend WithEvents CBoxCommitment As System.Windows.Forms.ComboBox
    Friend WithEvents LblComment As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents CkCommitment As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlIncompleteApplicationBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.pnlIncompleteApplicationDetails = New System.Windows.Forms.Panel
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.LblComment = New System.Windows.Forms.Label
        Me.CBoxCommitment = New System.Windows.Forms.ComboBox
        Me.LblCommitment = New System.Windows.Forms.Label
        Me.ReasonPanel = New System.Windows.Forms.Panel
        Me.clbIncompleteReason = New System.Windows.Forms.CheckedListBox
        Me.txtOther = New System.Windows.Forms.TextBox
        Me.lblOther = New System.Windows.Forms.Label
        Me.lblIncompleteApplication = New System.Windows.Forms.Label
        Me.chkIncompleteApp = New System.Windows.Forms.CheckBox
        Me.txtRequested = New System.Windows.Forms.TextBox
        Me.lblRequested = New System.Windows.Forms.Label
        Me.dtPickReceived = New System.Windows.Forms.DateTimePicker
        Me.lblReceived = New System.Windows.Forms.Label
        Me.CkCommitment = New System.Windows.Forms.CheckBox
        Me.pnlIncompleteApplicationBottom.SuspendLayout()
        Me.pnlIncompleteApplicationDetails.SuspendLayout()
        Me.ReasonPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlIncompleteApplicationBottom
        '
        Me.pnlIncompleteApplicationBottom.Controls.Add(Me.btnCancel)
        Me.pnlIncompleteApplicationBottom.Controls.Add(Me.btnOK)
        Me.pnlIncompleteApplicationBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlIncompleteApplicationBottom.Location = New System.Drawing.Point(0, 470)
        Me.pnlIncompleteApplicationBottom.Name = "pnlIncompleteApplicationBottom"
        Me.pnlIncompleteApplicationBottom.Size = New System.Drawing.Size(600, 40)
        Me.pnlIncompleteApplicationBottom.TabIndex = 4
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(303, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(223, 8)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        '
        'pnlIncompleteApplicationDetails
        '
        Me.pnlIncompleteApplicationDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.CkCommitment)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.txtComments)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.LblComment)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.CBoxCommitment)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.LblCommitment)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.ReasonPanel)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.chkIncompleteApp)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.txtRequested)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.lblRequested)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.dtPickReceived)
        Me.pnlIncompleteApplicationDetails.Controls.Add(Me.lblReceived)
        Me.pnlIncompleteApplicationDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlIncompleteApplicationDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlIncompleteApplicationDetails.Name = "pnlIncompleteApplicationDetails"
        Me.pnlIncompleteApplicationDetails.Size = New System.Drawing.Size(600, 470)
        Me.pnlIncompleteApplicationDetails.TabIndex = 1
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(80, 56)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(504, 64)
        Me.txtComments.TabIndex = 11
        Me.txtComments.Text = ""
        '
        'LblComment
        '
        Me.LblComment.Location = New System.Drawing.Point(8, 56)
        Me.LblComment.Name = "LblComment"
        Me.LblComment.Size = New System.Drawing.Size(80, 23)
        Me.LblComment.TabIndex = 10
        Me.LblComment.Text = "Comments :"
        '
        'CBoxCommitment
        '
        Me.CBoxCommitment.Enabled = False
        Me.CBoxCommitment.Location = New System.Drawing.Point(256, 32)
        Me.CBoxCommitment.Name = "CBoxCommitment"
        Me.CBoxCommitment.Size = New System.Drawing.Size(328, 21)
        Me.CBoxCommitment.TabIndex = 9
        '
        'LblCommitment
        '
        Me.LblCommitment.Location = New System.Drawing.Point(184, 32)
        Me.LblCommitment.Name = "LblCommitment"
        Me.LblCommitment.Size = New System.Drawing.Size(88, 23)
        Me.LblCommitment.TabIndex = 8
        Me.LblCommitment.Text = "Commitment:"
        '
        'ReasonPanel
        '
        Me.ReasonPanel.Controls.Add(Me.clbIncompleteReason)
        Me.ReasonPanel.Controls.Add(Me.txtOther)
        Me.ReasonPanel.Controls.Add(Me.lblOther)
        Me.ReasonPanel.Controls.Add(Me.lblIncompleteApplication)
        Me.ReasonPanel.Location = New System.Drawing.Point(0, 128)
        Me.ReasonPanel.Name = "ReasonPanel"
        Me.ReasonPanel.Size = New System.Drawing.Size(600, 344)
        Me.ReasonPanel.TabIndex = 7
        Me.ReasonPanel.Visible = False
        '
        'clbIncompleteReason
        '
        Me.clbIncompleteReason.CheckOnClick = True
        Me.clbIncompleteReason.Location = New System.Drawing.Point(16, 32)
        Me.clbIncompleteReason.Name = "clbIncompleteReason"
        Me.clbIncompleteReason.Size = New System.Drawing.Size(568, 259)
        Me.clbIncompleteReason.TabIndex = 8
        '
        'txtOther
        '
        Me.txtOther.Location = New System.Drawing.Point(80, 312)
        Me.txtOther.Name = "txtOther"
        Me.txtOther.Size = New System.Drawing.Size(385, 20)
        Me.txtOther.TabIndex = 9
        Me.txtOther.Text = ""
        '
        'lblOther
        '
        Me.lblOther.Location = New System.Drawing.Point(16, 312)
        Me.lblOther.Name = "lblOther"
        Me.lblOther.Size = New System.Drawing.Size(40, 23)
        Me.lblOther.TabIndex = 10
        Me.lblOther.Text = "Other:"
        '
        'lblIncompleteApplication
        '
        Me.lblIncompleteApplication.Location = New System.Drawing.Point(16, 8)
        Me.lblIncompleteApplication.Name = "lblIncompleteApplication"
        Me.lblIncompleteApplication.Size = New System.Drawing.Size(176, 17)
        Me.lblIncompleteApplication.TabIndex = 7
        Me.lblIncompleteApplication.Text = "Incomplete Application Reason(s)"
        '
        'chkIncompleteApp
        '
        Me.chkIncompleteApp.Location = New System.Drawing.Point(424, 0)
        Me.chkIncompleteApp.Name = "chkIncompleteApp"
        Me.chkIncompleteApp.Size = New System.Drawing.Size(168, 24)
        Me.chkIncompleteApp.TabIndex = 2
        Me.chkIncompleteApp.Text = "Incomplete Application"
        '
        'txtRequested
        '
        Me.txtRequested.Location = New System.Drawing.Point(79, 32)
        Me.txtRequested.Name = "txtRequested"
        Me.txtRequested.Size = New System.Drawing.Size(89, 20)
        Me.txtRequested.TabIndex = 1
        Me.txtRequested.Text = "0.00"
        Me.txtRequested.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblRequested
        '
        Me.lblRequested.Location = New System.Drawing.Point(8, 32)
        Me.lblRequested.Name = "lblRequested"
        Me.lblRequested.Size = New System.Drawing.Size(72, 23)
        Me.lblRequested.TabIndex = 2
        Me.lblRequested.Text = "Requested:"
        '
        'dtPickReceived
        '
        Me.dtPickReceived.Checked = False
        Me.dtPickReceived.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickReceived.Location = New System.Drawing.Point(79, 8)
        Me.dtPickReceived.Name = "dtPickReceived"
        Me.dtPickReceived.Size = New System.Drawing.Size(89, 20)
        Me.dtPickReceived.TabIndex = 0
        '
        'lblReceived
        '
        Me.lblReceived.Location = New System.Drawing.Point(8, 8)
        Me.lblReceived.Name = "lblReceived"
        Me.lblReceived.Size = New System.Drawing.Size(64, 23)
        Me.lblReceived.TabIndex = 0
        Me.lblReceived.Text = "Received:"
        '
        'CkCommitment
        '
        Me.CkCommitment.Location = New System.Drawing.Point(256, 0)
        Me.CkCommitment.Name = "CkCommitment"
        Me.CkCommitment.Size = New System.Drawing.Size(168, 24)
        Me.CkCommitment.TabIndex = 12
        Me.CkCommitment.Text = "Select Commitment?"
        '
        'IncompleteApplication
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(600, 510)
        Me.Controls.Add(Me.pnlIncompleteApplicationDetails)
        Me.Controls.Add(Me.pnlIncompleteApplicationBottom)
        Me.MaximizeBox = False
        Me.Name = "IncompleteApplication"
        Me.Text = "Reimbursement Request"
        Me.pnlIncompleteApplicationBottom.ResumeLayout(False)
        Me.pnlIncompleteApplicationDetails.ResumeLayout(False)
        Me.ReasonPanel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Local Variables "

    Friend FinancialEventID As Int64
    'Friend FinancialCommitmentID As Int64
    Friend FinancialReimbursementID As Int64

    Private bolFormatting As Boolean
    Private bolLoading As Boolean
    Private oFinancialEvent As New MUSTER.BusinessLogic.pFinancial
    Private oFinancialActivity As New MUSTER.BusinessLogic.pFinancialActivity
    Private oFinancialReimbursement As New MUSTER.BusinessLogic.pFinancialReimbursement

    Private OtherIndex As Int16
    Dim returnVal As String = String.Empty

    Friend CallingForm As Form
#End Region

#Region " Form Events "
    Private Sub IncompleteApplication_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bolLoading = True
        LoadCheckBoxList()
        bolLoading = False

        oFinancialEvent.Retrieve(FinancialEventID)
        SetUpCommitmentBoxes()

        oFinancialReimbursement.Retrieve(FinancialReimbursementID)

        txtComments.Text = oFinancialReimbursement.Comment

        LoadForm()

    End Sub

    Private Sub dtPickReceived_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickReceived.ValueChanged
        If bolLoading = True Then Exit Sub
        oFinancialReimbursement.ReceivedDate = dtPickReceived.Value.Date
    End Sub
    Private Sub chkIncompleteApp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncompleteApp.CheckedChanged
        If chkIncompleteApp.Checked Then
            Me.Height = 544

            Me.ReasonPanel.Visible = True
        Else
            Me.Height = 192
            Me.ReasonPanel.Visible = False

        End If
        oFinancialReimbursement.Incomplete = chkIncompleteApp.Checked
    End Sub
    Private Sub txtRequested_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRequested.TextChanged
        If bolLoading = True Or bolFormatting = True Then Exit Sub
        If txtRequested.Text = String.Empty Then Exit Sub

        If Not IsNumeric(txtRequested.Text) Then
            txtRequested.Text = "0.00"
        End If
        oFinancialReimbursement.RequestedAmount = txtRequested.Text
    End Sub
    Private Sub txtRequested_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRequested.LostFocus
        bolFormatting = True
        If txtRequested.Text = String.Empty Then
            oFinancialReimbursement.RequestedAmount = 0
            txtRequested.Text = "0.00"
        ElseIf IsNumeric(txtRequested.Text) Then
            txtRequested.Text = FormatNumber(txtRequested.Text, 2, TriState.False, TriState.False, TriState.True)
            oFinancialReimbursement.RequestedAmount = txtRequested.Text
        Else
            txtRequested.Text = String.Empty
            MsgBox("Numeric Data Only")
            oFinancialReimbursement.RequestedAmount = 0
        End If
        bolFormatting = False
    End Sub

    Private Sub clbIncompleteReason_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If clbIncompleteReason.GetItemChecked(OtherIndex) Then
            If Not OtherIndex = 12 Then txtOther.Text = ""
            txtOther.Visible = True
            lblOther.Visible = True
        Else
            txtOther.Text = ""
            txtOther.Visible = False
            lblOther.Visible = False
        End If
    End Sub


    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim bolGenerateIncompleteNotice As Boolean = False
        Dim oLetter As New Reg_Letters
        Dim strLongName As String
        Dim strShortName As String
        Dim strTemplate As String
        Dim oTechEvent As New MUSTER.BusinessLogic.pLustEvent
        Dim oFacility As New MUSTER.BusinessLogic.pFacility
        Dim oOwner As New MUSTER.BusinessLogic.pOwner
        Dim bolAllowNegativeRequest As Boolean = False

        '  Dim strSQL As String = "select count(Fin_Event_ID) from vFinancialCommitment_Grid where Fin_Event_ID = " + FinancialEventID.ToString + " and CommitmentID in (" + _
        '     "select fc.CommitmentID from tblfin_commitment fc where fc.fin_event_id = " + FinancialEventID.ToString + " and fc.case_letter in ('CR','IR')) " + _
        '     "and Balance <> '0.00' and len(Balance) > 0"

        Dim strSQL As String = "select count(Fin_Event_ID) from vFinancialCommitment_Grid where Fin_Event_ID = " + FinancialEventID.ToString + " and Balance <> '0.00' and len(Balance) > 0"
        If oOwner.RunSQLQuery(strSQL).Tables(0).Rows.Count > 0 Then
            If oOwner.RunSQLQuery(strSQL).Tables(0).Rows(0)(0) > 0 Then
                bolAllowNegativeRequest = True
            End If
        End If

        Try


            oFinancialReimbursement.Comment = Me.txtComments.Text

            If IsNumeric(txtRequested.Text) Then
                If txtRequested.Text > 0 Or (txtRequested.Text < 0 And bolAllowNegativeRequest) Then
                    If chkIncompleteApp.Checked Then
                        If GetIncompleteAppInfo() Is Nothing Then
                            MessageBox.Show("Please Select Atleast One Reason.")
                            Exit Sub
                        End If

                        oFinancialReimbursement.IncompleteReason = GetIncompleteAppInfo()
                        oFinancialReimbursement.IncompleteOther = txtOther.Text
                        If oFinancialReimbursement.id > 0 Then
                            If MsgBox("Generate New Notice Of Incomplete Application?", MsgBoxStyle.YesNo, "Incomplete Application") = MsgBoxResult.Yes Then
                                bolGenerateIncompleteNotice = True
                            End If
                        Else
                            bolGenerateIncompleteNotice = True
                        End If
                    Else
                        oFinancialReimbursement.Incomplete = False
                        oFinancialReimbursement.IncompleteOther = ""
                        oFinancialReimbursement.IncompleteReason = ""
                    End If

                    If oFinancialReimbursement.id <= 0 Then
                        oFinancialReimbursement.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oFinancialReimbursement.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oFinancialReimbursement.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    If Not CallingForm Is Nothing AndAlso Not CallingForm.Tag Is Nothing Then
                        CallingForm.Tag = "R" + oFinancialReimbursement.id.ToString
                    End If

                    MsgBox("Reimbursement Request Saved")

                    If bolGenerateIncompleteNotice Then
                        'Generate Incomplete Application Here
                        strLongName = "Notice Of Incomplete Application"
                        strShortName = "IncompleteApplication"
                        strTemplate = "NoticeofIncompleteApplicationTemplate.doc"
                        oTechEvent.Retrieve(oFinancialEvent.TecEventID)
                        oFacility.Retrieve(oOwner.OwnerInfo, oTechEvent.FacilityID, "SELF", "FACILITY")
                        oOwner.Retrieve(oFacility.OwnerID)


                        oLetter.GenerateFinancialLetter(oTechEvent.FacilityID, strLongName, strShortName, strLongName, strTemplate, oOwner, oFinancialEvent.TecEventID, oFinancialReimbursement.FinancialEventID, 0, oFinancialReimbursement.id, 0, 0, , oFinancialEvent.ID, oFinancialEvent.Sequence, UIUtilsGen.EntityTypes.FinancialEvent)
                    End If
                    Me.Close()
                Else
                    MsgBox("Invalid Requested Amount" + vbCrLf + "Must Be Greater Than Zero" + vbCrLf + "OR" + vbCrLf + "Less Than Zero if there are open Credit Request / Refund Commitment(s)")
                    txtRequested.Focus()
                End If
            Else
                MsgBox("Requested Amount Must Be Numeric")
                txtRequested.Focus()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

        oFinancialReimbursement.Reset()
        clbIncompleteReason.ClearSelected()
        Me.Close()

    End Sub
#End Region


    Private Sub SetUpCommitmentBoxes()

        With Me.CBoxCommitment

            Dim dtCommitment = oFinancialEvent.PopulateInvoiceCommitmentList(FinancialEventID)

            .DisplayMember = "CommitmentDesc"
            .ValueMember = "CommitmentID"

            .DataSource = dtCommitment

            If .Items.Count = 0 Then
                Me.CkCommitment.Enabled = False
            End If

        End With

    End Sub

    Private Sub LoadCheckBoxList()
        Dim dtTemp As DataTable
        Dim i As Int16
        dtTemp = oFinancialReimbursement.PopulateFinancialIncompleteAppReasons()

        clbIncompleteReason.DataSource = dtTemp
        clbIncompleteReason.DisplayMember = "Financial_Text"
        clbIncompleteReason.ValueMember = "Text_ID"
        OtherIndex = FindOtherInCheckBoxList("Other")
    End Sub

    Private Function FindOtherInCheckBoxList(ByVal SetItem As String) As Int64
        Dim i As Int16

        For i = 0 To clbIncompleteReason.Items.Count - 1
            If CType(CType(CType(clbIncompleteReason.Items(i), Object), System.Data.DataRowView).Row, System.Data.DataRow).ItemArray(1) = SetItem Then
                FindOtherInCheckBoxList = i
            End If
        Next
    End Function

    Private Sub SetCheckBoxList(ByVal SetItem As Int64)
        Dim i As Int16

        For i = 0 To clbIncompleteReason.Items.Count - 1
            If CType(CType(CType(clbIncompleteReason.Items(i), Object), System.Data.DataRowView).Row, System.Data.DataRow).ItemArray(0) = SetItem Then
                clbIncompleteReason.SetItemChecked(i, True)
            End If
        Next
    End Sub

    Private Sub LoadForm()
        Dim xArray As Array
        Dim i As Int16

        If FinancialReimbursementID = 0 Then
            Me.Height = 192
            txtOther.Text = ""
            txtOther.Visible = False
            lblOther.Visible = False
            dtPickReceived.Value = Now.Date
            txtRequested.Text = "0.00"
            chkIncompleteApp.Checked = False

            oFinancialReimbursement.ReceivedDate = Now.Date
            oFinancialReimbursement.RequestedAmount = 0
            oFinancialReimbursement.CommitmentID = 0
            oFinancialReimbursement.FinancialEventID = FinancialEventID
            oFinancialReimbursement.Incomplete = False
        Else
            dtPickReceived.Value = oFinancialReimbursement.ReceivedDate
            txtRequested.Text = FormatNumber(oFinancialReimbursement.RequestedAmount, 2, TriState.True, TriState.False, TriState.True)
            chkIncompleteApp.Checked = oFinancialReimbursement.Incomplete

            If oFinancialReimbursement.Incomplete Then
                xArray = oFinancialReimbursement.IncompleteReason.Split(",")
                For i = 0 To xArray.Length - 1
                    SetCheckBoxList(xArray(i))
                Next
            Else

                Me.Height = 192
            End If


            txtOther.Text = oFinancialReimbursement.IncompleteOther
            If clbIncompleteReason.GetItemChecked(OtherIndex) Then
                txtOther.Visible = True
                lblOther.Visible = True
            Else
                txtOther.Text = ""
                txtOther.Visible = False
                lblOther.Visible = False
            End If

        End If

    End Sub

    Private Function GetIncompleteAppInfo() As String
        Dim i As Integer
        Dim strComma As String
        Dim strReturn As String
        strComma = ""
        Try

            For i = 0 To clbIncompleteReason.CheckedItems.Count - 1
                strReturn &= strComma & CType(CType(CType(clbIncompleteReason.CheckedItems(i), Object), System.Data.DataRowView).Row, System.Data.DataRow).ItemArray(0)
                strComma = ","
            Next

            Return strReturn
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally

        End Try


    End Function


    Private Sub CBoxCommitment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBoxCommitment.SelectedIndexChanged

        If Me.CkCommitment.Checked Then
            oFinancialReimbursement.CommitmentID = CBoxCommitment.SelectedValue
        End If


    End Sub

    Private Sub CkCommitment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CkCommitment.CheckedChanged

        If Me.CkCommitment.Checked Then
            Me.CBoxCommitment.Enabled = True

            If Not oFinancialReimbursement.CommitmentID = Nothing Then
                Me.CBoxCommitment.SelectedValue = oFinancialReimbursement.CommitmentID
            Else
                Me.CBoxCommitment.SelectedIndex = 0
            End If
        Else

            Me.CBoxCommitment.Enabled = False
            oFinancialReimbursement.CommitmentID = Nothing

        End If




    End Sub
End Class
