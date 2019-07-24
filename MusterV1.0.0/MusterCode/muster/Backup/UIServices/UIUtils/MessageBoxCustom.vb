Public Class MessageBoxCustom
    Private Shared retValue As System.Windows.Forms.DialogResult
    Private Shared frm As New Form
    Private Shared txtBox As New TextBox
    Private Shared pnlTop As New Panel
    Private Shared pnlBottom As New Panel
    Private Shared dt As DataTable
    Private Shared WithEvents btnAbort As New System.Windows.Forms.Button
    Private Shared WithEvents btnRetry As New System.Windows.Forms.Button
    Private Shared WithEvents btnIgnore As New System.Windows.Forms.Button
    Private Shared WithEvents btnOK As New System.Windows.Forms.Button
    Private Shared WithEvents btnCancel As New System.Windows.Forms.Button
    Private Shared WithEvents btnYes As New System.Windows.Forms.Button
    Private Shared WithEvents btnNo As New System.Windows.Forms.Button
    Private Shared WithEvents btnOKOnly As New System.Windows.Forms.Button

    Friend Shared Function Show(ByVal text As String, Optional ByVal caption As String = "", _
                                Optional ByVal buttons As System.Windows.Forms.MessageBoxButtons = MessageBoxButtons.OK, _
                                Optional ByVal icon As System.Windows.Forms.MessageBoxIcon = MessageBoxIcon.None, _
                                Optional ByVal textAlign As System.Windows.Forms.HorizontalAlignment = HorizontalAlignment.Center, _
                                Optional ByVal txtBoxBackColor As System.Drawing.KnownColor = KnownColor.Control, _
                                Optional ByVal txtBoxForeColor As System.Drawing.KnownColor = KnownColor.Black) As System.Windows.Forms.DialogResult
        Initialize()

        'Dim panelIndex As Integer = frm.Controls.IndexOf(pnlTop)
        'Dim index As Integer = frm.Controls(panelIndex).Controls.IndexOf(txtBox)
        txtBox.BackColor = Color.FromKnownColor(txtBoxBackColor)
        txtBox.Text = text
        txtBox.TextAlign = textAlign
        txtBox.ForeColor = Color.FromKnownColor(txtBoxForeColor)

        Select Case buttons
            Case MessageBoxButtons.AbortRetryIgnore
                btnAbort.Visible = True
                btnRetry.Visible = True
                btnIgnore.Visible = True
                btnAbort.Select()
                retValue = DialogResult.Abort
                'panelIndex = frm.Controls.IndexOf(pnlBottom)
                'index = frm.Controls(panelIndex).Controls.IndexOf(btnAbort)
            Case MessageBoxButtons.OK
                btnOKOnly.Visible = True
                btnOKOnly.Select()
                'panelIndex = frm.Controls.IndexOf(pnlBottom)
                'index = frm.Controls(panelIndex).Controls.IndexOf(btnOKOnly)
                retValue = DialogResult.OK
            Case MessageBoxButtons.OKCancel
                btnOK.Visible = True
                btnCancel.Visible = True
                btnOK.Select()
                'panelIndex = frm.Controls.IndexOf(pnlBottom)
                'index = frm.Controls(panelIndex).Controls.IndexOf(btnOK)
                retValue = DialogResult.Cancel
            Case MessageBoxButtons.RetryCancel
                btnRetry.Visible = True
                btnCancel.Visible = True
                btnRetry.Location = btnOK.Location
                btnRetry.Select()
                'panelIndex = frm.Controls.IndexOf(pnlBottom)
                'index = frm.Controls(panelIndex).Controls.IndexOf(btnRetry)
                retValue = DialogResult.Cancel
            Case MessageBoxButtons.YesNo
                btnYes.Visible = True
                btnNo.Visible = True
                btnYes.Select()
                'panelIndex = frm.Controls.IndexOf(pnlBottom)
                'index = frm.Controls(panelIndex).Controls.IndexOf(btnYes)
                retValue = DialogResult.No
            Case MessageBoxButtons.YesNoCancel
                btnYes.Visible = True
                btnNo.Visible = True
                btnCancel.Visible = True
                btnYes.Location = btnAbort.Location
                btnNo.Location = btnRetry.Location
                btnCancel.Location = btnIgnore.Location
                btnYes.Select()
                'panelIndex = frm.Controls.IndexOf(pnlBottom)
                'index = frm.Controls(panelIndex).Controls.IndexOf(btnYes)
                retValue = DialogResult.Cancel
        End Select

        If caption = String.Empty Then
            frm.Text = System.AppDomain.CurrentDomain.FriendlyName
        Else
            frm.Text = caption
        End If

        'frm.Controls(panelIndex).Controls(index).Focus()
        frm.ShowDialog()

        Return retValue
    End Function

    Private Shared Sub Initialize()
        frm = New Form

        txtBox = New System.Windows.Forms.TextBox
        pnlTop = New System.Windows.Forms.Panel
        pnlBottom = New System.Windows.Forms.Panel
        btnAbort = New System.Windows.Forms.Button
        btnRetry = New System.Windows.Forms.Button
        btnIgnore = New System.Windows.Forms.Button
        btnOK = New System.Windows.Forms.Button
        btnCancel = New System.Windows.Forms.Button
        btnYes = New System.Windows.Forms.Button
        btnNo = New System.Windows.Forms.Button
        btnOKOnly = New System.Windows.Forms.Button
        pnlTop.SuspendLayout()
        frm.SuspendLayout()
        '
        'txtBox
        '
        txtBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        txtBox.Dock = System.Windows.Forms.DockStyle.Fill
        txtBox.ReadOnly = True
        txtBox.Location = New System.Drawing.Point(0, 0)
        txtBox.Multiline = True
        txtBox.Name = "txtBox"
        txtBox.ScrollBars = System.Windows.Forms.ScrollBars.Both
        txtBox.Size = New System.Drawing.Size(292, 233)
        txtBox.TabIndex = 8
        txtBox.Text = ""
        txtBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'pnlTop
        '
        pnlTop.Controls.Add(txtBox)
        pnlTop.Dock = System.Windows.Forms.DockStyle.Fill
        pnlTop.Location = New System.Drawing.Point(0, 0)
        pnlTop.Name = "pnlTop"
        pnlTop.Size = New System.Drawing.Size(292, 233)
        pnlTop.TabIndex = 0
        '
        'pnlBottom
        '
        pnlBottom.Controls.Add(btnAbort)
        pnlBottom.Controls.Add(btnRetry)
        pnlBottom.Controls.Add(btnIgnore)
        pnlBottom.Controls.Add(btnOK)
        pnlBottom.Controls.Add(btnCancel)
        pnlBottom.Controls.Add(btnYes)
        pnlBottom.Controls.Add(btnNo)
        pnlBottom.Controls.Add(btnOKOnly)
        pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        pnlBottom.Location = New System.Drawing.Point(0, 201)
        pnlBottom.Name = "pnlBottom"
        pnlBottom.Size = New System.Drawing.Size(292, 40)
        pnlBottom.TabIndex = 1
        '
        'btnAbort
        '
        btnAbort.Location = New System.Drawing.Point(29, 8)
        btnAbort.Name = "btnAbort"
        btnAbort.TabIndex = 0
        btnAbort.Text = "Abort"
        btnAbort.Visible = False
        '
        'btnRetry
        '
        btnRetry.Location = New System.Drawing.Point(109, 8)
        btnRetry.Name = "btnRetry"
        btnRetry.TabIndex = 1
        btnRetry.Text = "Retry"
        btnRetry.Visible = False
        '
        'btnIgnore
        '
        btnIgnore.Location = New System.Drawing.Point(189, 8)
        btnIgnore.Name = "btnIgnore"
        btnIgnore.TabIndex = 2
        btnIgnore.Text = "Ignore"
        btnIgnore.Visible = False
        '
        'btnOK
        '
        btnOK.Location = New System.Drawing.Point(69, 8)
        btnOK.Name = "btnOK"
        btnOK.TabIndex = 3
        btnOK.Text = "OK"
        btnOK.Visible = False
        '
        'btnCancel
        '
        btnCancel.Location = New System.Drawing.Point(149, 8)
        btnCancel.Name = "btnCancel"
        btnCancel.TabIndex = 4
        btnCancel.Text = "Cancel"
        btnCancel.Visible = False
        '
        'btnYes
        '
        btnYes.Location = New System.Drawing.Point(69, 8)
        btnYes.Name = "btnYes"
        btnYes.TabIndex = 5
        btnYes.Text = "Yes"
        btnYes.Visible = False
        '
        'btnNo
        '
        btnNo.Location = New System.Drawing.Point(149, 8)
        btnNo.Name = "btnNo"
        btnNo.TabIndex = 6
        btnNo.Text = "No"
        btnNo.Visible = False
        '
        'btnOKOnly
        '
        btnOKOnly.Location = New System.Drawing.Point(109, 8)
        btnOKOnly.Name = "btnOKOnly"
        btnOKOnly.TabIndex = 7
        btnOKOnly.Text = "OK"
        btnOKOnly.Visible = False
        '
        'frm
        '
        frm.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        frm.AutoScroll = True
        frm.ControlBox = False
        'frm.MinimizeBox = False
        'frm.MaximizeBox = True
        frm.ClientSize = New System.Drawing.Size(292, 273)
        frm.Controls.Add(pnlTop)
        frm.Controls.Add(pnlBottom)
        frm.Name = "frm"
        frm.Text = "frm"
        frm.ShowInTaskbar = True
        frm.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        'CType(ug, System.ComponentModel.ISupportInitialize).EndInit()
        pnlTop.ResumeLayout(False)
        frm.ResumeLayout(False)

        dt = New DataTable
        dt.Columns.Add("DISPLAYSTRING", GetType(String))
    End Sub

    Private Shared Sub btnAbort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAbort.Click
        retValue = DialogResult.Abort
        frm.Close()
    End Sub

    Private Shared Sub btnRetry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRetry.Click
        retValue = DialogResult.Retry
        frm.Close()
    End Sub

    Private Shared Sub btnIgnore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIgnore.Click
        retValue = DialogResult.Ignore
        frm.Close()
    End Sub

    Private Shared Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        retValue = DialogResult.OK
        frm.Close()
    End Sub

    Private Shared Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        retValue = DialogResult.Cancel
        frm.Close()
    End Sub

    Private Shared Sub btnYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnYes.Click
        retValue = DialogResult.Yes
        frm.Close()
    End Sub

    Private Shared Sub btnNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNo.Click
        retValue = DialogResult.No
        frm.Close()
    End Sub

    Private Shared Sub btnOKOnly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOKOnly.Click
        retValue = DialogResult.OK
        frm.Close()
    End Sub

End Class
