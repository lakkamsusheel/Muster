Public Class ChangePassword
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.ChangePassword
    '   The interface to change user logon passwords for the MUSTER application.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      ??/??/??    Original class definition.
    '  1.1        JC      12/31/04    Changed over to Muster.BussinessLogic.pUser
    '-------------------------------------------------------------------------------
    Inherits System.Windows.Forms.Form
    Dim LocalUser As Muster.BusinessLogic.pUser
    Friend Event PasswordChanged()
    Friend Event PasswordClosed()


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
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents lblNewPassword As System.Windows.Forms.Label
    Friend WithEvents lblOldPassword As System.Windows.Forms.Label
    Friend WithEvents txtNewPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtOldPassword As System.Windows.Forms.TextBox
    Friend WithEvents lblConfirmNewPwd As System.Windows.Forms.Label
    Friend WithEvents txtConfirmNewPwd As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.lblNewPassword = New System.Windows.Forms.Label
        Me.lblOldPassword = New System.Windows.Forms.Label
        Me.txtNewPassword = New System.Windows.Forms.TextBox
        Me.txtOldPassword = New System.Windows.Forms.TextBox
        Me.lblConfirmNewPwd = New System.Windows.Forms.Label
        Me.txtConfirmNewPwd = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(136, 120)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 23)
        Me.btnCancel.TabIndex = 141
        Me.btnCancel.Text = "Cancel"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(64, 120)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(64, 23)
        Me.btnOK.TabIndex = 140
        Me.btnOK.Text = "OK"
        '
        'lblNewPassword
        '
        Me.lblNewPassword.Location = New System.Drawing.Point(16, 48)
        Me.lblNewPassword.Name = "lblNewPassword"
        Me.lblNewPassword.Size = New System.Drawing.Size(80, 16)
        Me.lblNewPassword.TabIndex = 139
        Me.lblNewPassword.Text = "New Password"
        Me.lblNewPassword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblOldPassword
        '
        Me.lblOldPassword.Location = New System.Drawing.Point(16, 24)
        Me.lblOldPassword.Name = "lblOldPassword"
        Me.lblOldPassword.Size = New System.Drawing.Size(80, 16)
        Me.lblOldPassword.TabIndex = 138
        Me.lblOldPassword.Text = "Old Password"
        Me.lblOldPassword.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtNewPassword
        '
        Me.txtNewPassword.Location = New System.Drawing.Point(104, 48)
        Me.txtNewPassword.Name = "txtNewPassword"
        Me.txtNewPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtNewPassword.Size = New System.Drawing.Size(128, 20)
        Me.txtNewPassword.TabIndex = 137
        Me.txtNewPassword.Text = ""
        '
        'txtOldPassword
        '
        Me.txtOldPassword.Location = New System.Drawing.Point(104, 24)
        Me.txtOldPassword.Name = "txtOldPassword"
        Me.txtOldPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtOldPassword.Size = New System.Drawing.Size(128, 20)
        Me.txtOldPassword.TabIndex = 136
        Me.txtOldPassword.Text = ""
        '
        'lblConfirmNewPwd
        '
        Me.lblConfirmNewPwd.Location = New System.Drawing.Point(16, 80)
        Me.lblConfirmNewPwd.Name = "lblConfirmNewPwd"
        Me.lblConfirmNewPwd.Size = New System.Drawing.Size(72, 24)
        Me.lblConfirmNewPwd.TabIndex = 143
        Me.lblConfirmNewPwd.Text = "Confirm New Password"
        Me.lblConfirmNewPwd.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtConfirmNewPwd
        '
        Me.txtConfirmNewPwd.Location = New System.Drawing.Point(104, 80)
        Me.txtConfirmNewPwd.Name = "txtConfirmNewPwd"
        Me.txtConfirmNewPwd.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtConfirmNewPwd.Size = New System.Drawing.Size(128, 20)
        Me.txtConfirmNewPwd.TabIndex = 138
        Me.txtConfirmNewPwd.Text = ""
        '
        'ChangePassword
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(264, 158)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblConfirmNewPwd)
        Me.Controls.Add(Me.txtConfirmNewPwd)
        Me.Controls.Add(Me.txtNewPassword)
        Me.Controls.Add(Me.txtOldPassword)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.lblNewPassword)
        Me.Controls.Add(Me.lblOldPassword)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ChangePassword"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change Password"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub SetUser(ByRef oUser As Muster.BusinessLogic.pUser)
        LocalUser = oUser
    End Sub
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If txtOldPassword.Text <> "" And txtNewPassword.Text <> "" And txtConfirmNewPwd.Text <> "" Then
            Try
                LocalUser.Password = txtOldPassword.Text
                If LocalUser.VerifyPassword = True Then

                    If Not txtNewPassword.Text <> txtOldPassword.Text Then
                        MsgBox("Duplicate Entry. Please try again!", MsgBoxStyle.Critical & MsgBoxStyle.OKOnly, "Duplicate Entry")
                        Exit Sub
                    End If

                    'for Ciber Bug ID : 561
                    If txtNewPassword.Text.ToLower.IndexOf("password") >= 0 Then
                        MsgBox("Please try some other Password other than 'password'.")
                        Exit Sub
                    End If

                    If txtNewPassword.Text <> Me.txtConfirmNewPwd.Text Then
                        MsgBox("New Passwords does not match. Please try again!")
                        Exit Sub
                    End If

                    LocalUser.Password = txtNewPassword.Text
                    MsgBox("Password Reset")
                    RaiseEvent PasswordChanged()
                    Me.Close()

                Else
                    MsgBox("Old Password was invalid. Please try again!")
                End If
            Catch ex As Exception
                Dim erbox As New ErrorReport(ex)
                erbox.ShowDialog()
            End Try
        Else
            MsgBox("You did not enter anything.")
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'Added to Close the Application.
        Me.Close()
        RaiseEvent PasswordClosed()

    End Sub
End Class
