Public Class CandEStatusSelection
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
    Friend WithEvents BtnNew As System.Windows.Forms.Button
    Friend WithEvents ButtonAgreedOrder As System.Windows.Forms.Button
    Friend WithEvents BtnRedTagPending As System.Windows.Forms.Button
    Friend WithEvents btnRedTag As System.Windows.Forms.Button
    Friend WithEvents btn2NDNotice As System.Windows.Forms.Button
    Friend WithEvents BtnNFA As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.BtnNew = New System.Windows.Forms.Button
        Me.ButtonAgreedOrder = New System.Windows.Forms.Button
        Me.BtnRedTagPending = New System.Windows.Forms.Button
        Me.btnRedTag = New System.Windows.Forms.Button
        Me.btn2NDNotice = New System.Windows.Forms.Button
        Me.BtnNFA = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'BtnNew
        '
        Me.BtnNew.Location = New System.Drawing.Point(8, 8)
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.Size = New System.Drawing.Size(104, 23)
        Me.BtnNew.TabIndex = 0
        Me.BtnNew.Tag = "0"
        Me.BtnNew.Text = "New"
        '
        'ButtonAgreedOrder
        '
        Me.ButtonAgreedOrder.Location = New System.Drawing.Point(8, 40)
        Me.ButtonAgreedOrder.Name = "ButtonAgreedOrder"
        Me.ButtonAgreedOrder.Size = New System.Drawing.Size(104, 23)
        Me.ButtonAgreedOrder.TabIndex = 1
        Me.ButtonAgreedOrder.Tag = "1"
        Me.ButtonAgreedOrder.Text = "Agreed Order"
        '
        'BtnRedTagPending
        '
        Me.BtnRedTagPending.Location = New System.Drawing.Point(136, 8)
        Me.BtnRedTagPending.Name = "BtnRedTagPending"
        Me.BtnRedTagPending.Size = New System.Drawing.Size(104, 23)
        Me.BtnRedTagPending.TabIndex = 2
        Me.BtnRedTagPending.Tag = "2"
        Me.BtnRedTagPending.Text = "RedTag Pending"
        '
        'btnRedTag
        '
        Me.btnRedTag.Location = New System.Drawing.Point(136, 40)
        Me.btnRedTag.Name = "btnRedTag"
        Me.btnRedTag.Size = New System.Drawing.Size(104, 23)
        Me.btnRedTag.TabIndex = 3
        Me.btnRedTag.Tag = "3"
        Me.btnRedTag.Text = "Red Tag"
        '
        'btn2NDNotice
        '
        Me.btn2NDNotice.Location = New System.Drawing.Point(256, 8)
        Me.btn2NDNotice.Name = "btn2NDNotice"
        Me.btn2NDNotice.Size = New System.Drawing.Size(112, 23)
        Me.btn2NDNotice.TabIndex = 4
        Me.btn2NDNotice.Tag = "4"
        Me.btn2NDNotice.Text = "2nd Notice"
        '
        'BtnNFA
        '
        Me.BtnNFA.Location = New System.Drawing.Point(256, 40)
        Me.BtnNFA.Name = "BtnNFA"
        Me.BtnNFA.Size = New System.Drawing.Size(112, 23)
        Me.BtnNFA.TabIndex = 5
        Me.BtnNFA.Tag = "5"
        Me.BtnNFA.Text = "NFA"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(256, 72)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(112, 24)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Tag = "6"
        Me.btnCancel.Text = "Cancel"
        '
        'CandEStatusSelection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(384, 101)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.BtnNFA)
        Me.Controls.Add(Me.btn2NDNotice)
        Me.Controls.Add(Me.btnRedTag)
        Me.Controls.Add(Me.BtnRedTagPending)
        Me.Controls.Add(Me.ButtonAgreedOrder)
        Me.Controls.Add(Me.BtnNew)
        Me.MaximumSize = New System.Drawing.Size(392, 128)
        Me.MinimumSize = New System.Drawing.Size(392, 128)
        Me.Name = "CandEStatusSelection"
        Me.Text = "Select which OCE Status to Escalate"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "public members"
    Public SelectedButton As Button = Nothing

#End Region


#Region "event"

    Public Sub cancelClicked(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
        SelectedButton = Nothing
        Me.Close()
    End Sub

    Public Sub btnClicked(ByVal sender As Object, ByVal e As EventArgs) Handles btn2NDNotice.Click, BtnNew.Click, BtnNFA.Click, btnRedTag.Click, BtnRedTagPending.Click, ButtonAgreedOrder.Click
        Me.DialogResult = DialogResult.OK
        SelectedButton = sender
        Me.Close()
    End Sub

#End Region

End Class
