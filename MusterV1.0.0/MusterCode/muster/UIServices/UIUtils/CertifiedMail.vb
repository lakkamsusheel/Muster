Public Class CertifiedMail
    Inherits System.Windows.Forms.Form

    Friend strTxtNo1, strTxtNo2, strTxtNo3, strTxtNo4, strTxtNo5 As String
    Public Event evtCertifiedMail(ByVal strCertMail As String)
    Private bolAllowClose As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        strTxtNo1 = String.Empty
        strTxtNo2 = String.Empty
        strTxtNo3 = String.Empty
        strTxtNo4 = String.Empty
        strTxtNo5 = String.Empty
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
    Friend WithEvents lbl As System.Windows.Forms.Label
    Friend WithEvents txtNo2 As System.Windows.Forms.TextBox
    Friend WithEvents txtNo3 As System.Windows.Forms.TextBox
    Friend WithEvents txtNo4 As System.Windows.Forms.TextBox
    Friend WithEvents txtNo5 As System.Windows.Forms.TextBox
    Friend WithEvents txtNo1 As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lbl = New System.Windows.Forms.Label
        Me.txtNo2 = New System.Windows.Forms.TextBox
        Me.txtNo3 = New System.Windows.Forms.TextBox
        Me.txtNo4 = New System.Windows.Forms.TextBox
        Me.txtNo5 = New System.Windows.Forms.TextBox
        Me.txtNo1 = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lbl
        '
        Me.lbl.Location = New System.Drawing.Point(8, 8)
        Me.lbl.Name = "lbl"
        Me.lbl.Size = New System.Drawing.Size(216, 23)
        Me.lbl.TabIndex = 7
        Me.lbl.Text = "Please provide CERTIFIED MAIL Number"
        Me.lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtNo2
        '
        Me.txtNo2.Location = New System.Drawing.Point(60, 40)
        Me.txtNo2.Name = "txtNo2"
        Me.txtNo2.Size = New System.Drawing.Size(32, 20)
        Me.txtNo2.TabIndex = 1
        Me.txtNo2.Text = ""
        '
        'txtNo3
        '
        Me.txtNo3.Location = New System.Drawing.Point(100, 40)
        Me.txtNo3.Name = "txtNo3"
        Me.txtNo3.Size = New System.Drawing.Size(32, 20)
        Me.txtNo3.TabIndex = 2
        Me.txtNo3.Text = ""
        '
        'txtNo4
        '
        Me.txtNo4.Location = New System.Drawing.Point(139, 40)
        Me.txtNo4.Name = "txtNo4"
        Me.txtNo4.Size = New System.Drawing.Size(32, 20)
        Me.txtNo4.TabIndex = 3
        Me.txtNo4.Text = ""
        '
        'txtNo5
        '
        Me.txtNo5.Location = New System.Drawing.Point(178, 40)
        Me.txtNo5.Name = "txtNo5"
        Me.txtNo5.Size = New System.Drawing.Size(32, 20)
        Me.txtNo5.TabIndex = 4
        Me.txtNo5.Text = ""
        '
        'txtNo1
        '
        Me.txtNo1.Location = New System.Drawing.Point(21, 40)
        Me.txtNo1.Name = "txtNo1"
        Me.txtNo1.Size = New System.Drawing.Size(32, 20)
        Me.txtNo1.TabIndex = 0
        Me.txtNo1.Text = ""
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(35, 72)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(123, 72)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        '
        'CertifiedMail
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(232, 109)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.lbl)
        Me.Controls.Add(Me.txtNo2)
        Me.Controls.Add(Me.txtNo3)
        Me.Controls.Add(Me.txtNo4)
        Me.Controls.Add(Me.txtNo5)
        Me.Controls.Add(Me.txtNo1)
        Me.Controls.Add(Me.btnCancel)
        Me.Name = "CertifiedMail"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Certified Mail"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtNo1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo1.TextChanged
        strTxtNo1 = txtNo1.Text
        If txtNo1.Text.Length = 4 Then
            txtNo2.Focus()
        End If
    End Sub

    Private Sub txtNo2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo2.TextChanged
        strTxtNo2 = txtNo2.Text
        If txtNo2.Text.Length = 4 Then
            txtNo3.Focus()
        End If
    End Sub

    Private Sub txtNo3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo3.TextChanged
        strTxtNo3 = txtNo3.Text
        If txtNo3.Text.Length = 4 Then
            txtNo4.Focus()
        End If
    End Sub

    Private Sub txtNo4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo4.TextChanged
        strTxtNo4 = txtNo4.Text
        If txtNo4.Text.Length = 4 Then
            txtNo5.Focus()
        End If
    End Sub

    Private Sub txtNo5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo5.TextChanged
        strTxtNo5 = txtNo5.Text
        If txtNo5.Text.Length = 4 Then
            btnOK.Focus()
        End If
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        RaiseEvent evtCertifiedMail(strTxtNo1 + " " + strTxtNo2 + " " + strTxtNo3 + " " + strTxtNo4 + " " + strTxtNo5)
        bolAllowClose = True
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        bolAllowClose = True
        Me.Close()
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        If Not bolAllowClose Then
            If MsgBox("Do you want to cancel entering CERTIFIED MAIL Number?", MsgBoxStyle.YesNo, "Cancel Certified Mail Entry") = MsgBoxResult.No Then
                e.Cancel = True
            End If
        End If
    End Sub

End Class
