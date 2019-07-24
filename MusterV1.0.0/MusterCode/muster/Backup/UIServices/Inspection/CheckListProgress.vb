Public Class CheckListProgress
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
    Friend WithEvents pgBar As System.Windows.Forms.ProgressBar
    Friend WithEvents lblHeader As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pgBar = New System.Windows.Forms.ProgressBar
        Me.lblHeader = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'pgBar
        '
        Me.pgBar.Cursor = System.Windows.Forms.Cursors.Default
        Me.pgBar.Location = New System.Drawing.Point(0, 30)
        Me.pgBar.Name = "pgBar"
        Me.pgBar.Size = New System.Drawing.Size(184, 30)
        Me.pgBar.TabIndex = 2
        '
        'lblHeader
        '
        Me.lblHeader.Location = New System.Drawing.Point(0, 0)
        Me.lblHeader.Name = "lblHeader"
        Me.lblHeader.Size = New System.Drawing.Size(184, 25)
        Me.lblHeader.TabIndex = 3
        Me.lblHeader.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CheckListProgress
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(184, 60)
        Me.Controls.Add(Me.lblHeader)
        Me.Controls.Add(Me.pgBar)
        Me.Cursor = System.Windows.Forms.Cursors.AppStarting
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "CheckListProgress"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CheckListProgress"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public Property HeaderText() As String
        Get
            Return lblHeader.Text()
        End Get
        Set(ByVal Value As String)
            lblHeader.Text = Value
        End Set
    End Property
    Public Property ProgressBarValue() As Integer
        Get
            Return pgBar.Value
        End Get
        Set(ByVal Value As Integer)
            If Value >= pgBar.Maximum Then
                Me.Close()
            Else
                pgBar.Value = Value
            End If
        End Set
    End Property
    Public ReadOnly Property ProgressBarMax() As Integer
        Get
            Return pgBar.Maximum
        End Get
    End Property
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub CheckListProgress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        pgBar.Maximum = 100
        pgBar.Value = 0
        Me.TopMost = True
    End Sub
End Class
