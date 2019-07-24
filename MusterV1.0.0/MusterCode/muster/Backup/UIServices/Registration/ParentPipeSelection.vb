Public Class ParentPipeSelection
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RBtnNo As System.Windows.Forms.RadioButton
    Friend WithEvents RBtnYes As System.Windows.Forms.RadioButton
    Friend WithEvents LstParentPipes As System.Windows.Forms.ComboBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.RBtnNo = New System.Windows.Forms.RadioButton
        Me.RBtnYes = New System.Windows.Forms.RadioButton
        Me.LstParentPipes = New System.Windows.Forms.ComboBox
        Me.OKButton = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(280, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Is this going to be a pipe extension ?"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'RBtnNo
        '
        Me.RBtnNo.Checked = True
        Me.RBtnNo.Location = New System.Drawing.Point(240, 8)
        Me.RBtnNo.Name = "RBtnNo"
        Me.RBtnNo.TabIndex = 2
        Me.RBtnNo.TabStop = True
        Me.RBtnNo.Text = "No"
        '
        'RBtnYes
        '
        Me.RBtnYes.Location = New System.Drawing.Point(304, 8)
        Me.RBtnYes.Name = "RBtnYes"
        Me.RBtnYes.TabIndex = 3
        Me.RBtnYes.Text = "Yes"
        '
        'LstParentPipes
        '
        Me.LstParentPipes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.LstParentPipes.Enabled = False
        Me.LstParentPipes.Location = New System.Drawing.Point(16, 56)
        Me.LstParentPipes.Name = "LstParentPipes"
        Me.LstParentPipes.Size = New System.Drawing.Size(416, 21)
        Me.LstParentPipes.TabIndex = 5
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(360, 144)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "OK"
        '
        'ParentPipeSelection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(464, 174)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.LstParentPipes)
        Me.Controls.Add(Me.RBtnYes)
        Me.Controls.Add(Me.RBtnNo)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ParentPipeSelection"
        Me.Text = "Adding New Pipe or Pipe Extension"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Public Accessors"

    Public ParentPipeID As Integer = 0
    Public IsExtension As Boolean = False
    Public ParentPipeData As DataTable

#End Region

#Region "Form Events"

    Private Sub form_load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        'Set up event handling
        AddHandler RBtnYes.CheckedChanged, AddressOf RadioButton_CheckedChanged
        AddHandler RBtnNo.CheckedChanged, AddressOf RadioButton_CheckedChanged

        'Load List
        With LstParentPipes

            .ValueMember = "PIPE ID"
            .DisplayMember = "DESC"
            .DataSource = ParentPipeData

        End With




    End Sub


    Private Sub LstParentPipes_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LstParentPipes.SelectedIndexChanged

        ParentPipeID = LstParentPipes.SelectedValue

    End Sub

    Private Sub RadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If RBtnYes.Checked Then
            LstParentPipes.Enabled = True
            IsExtension = True
        Else
            LstParentPipes.Enabled = False
            IsExtension = False
        End If

    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click

        Me.DialogResult = DialogResult.OK

        Me.Close()

    End Sub

#End Region

End Class
