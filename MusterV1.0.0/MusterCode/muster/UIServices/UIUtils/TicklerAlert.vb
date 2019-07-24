Option Explicit On 
Option Strict On


Public Class TicklerAlert
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
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblMessage = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblMessage
        '
        Me.lblMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblMessage.Location = New System.Drawing.Point(24, 16)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(248, 23)
        Me.lblMessage.TabIndex = 0
        Me.lblMessage.Text = "You have new messages!"
        Me.lblMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TicklerAlert
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.ClientSize = New System.Drawing.Size(290, 62)
        Me.Controls.Add(Me.lblMessage)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximumSize = New System.Drawing.Size(298, 88)
        Me.MinimumSize = New System.Drawing.Size(298, 88)
        Me.Name = "TicklerAlert"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "MUSTER Tickler Alert"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "members"

    Private WithEvents _timer As System.Timers.Timer
    Private _container As MusterContainer
    Private _bit As Boolean = False
    Private _forced As Boolean = False




#End Region


#Region "Construct"

    Sub New(ByVal mdiContainer As MusterContainer)

        Me.New()

        _container = mdiContainer

    End Sub

#End Region


#Region "Methods"



    Public Sub StartAlert(ByVal forced As Boolean)

        _forced = forced

        _bit = False
        'Opacity = 0


        If _forced Then

            TextClicked(lblMessage, Nothing)

            Exit Sub

        End If

        'TopMost = True

        With _container

            'Location = New Point(.Width - Width - 20, .Height - Height - 20)


        End With

        If Not Visible Then

            Show()

            _timer = New System.Timers.Timer(500)

            _timer.Start()


            Refresh()

        End If



    End Sub



    Sub TextClicked(ByVal sender As Object, ByVal e As EventArgs) Handles lblMessage.Click

        _container.UpdateTicklerButton()

        _container.TicklerScreen.ShowActualForm(Me)

    End Sub

    Sub TimerHit(ByVal sender As Object, ByVal e As Timers.ElapsedEventArgs) Handles _timer.Elapsed

        If Not _bit Then

            'Opacity += 0.25

            If Opacity >= 1 Then
                Opacity = 1

                _bit = True
                _timer.Interval = 300
            End If
        Else
            'Opacity -= 0.08

            If Opacity <= 0 Then
                Opacity = 0

                _timer.Stop()

                _timer.Dispose()

                Me.Close()


            End If

        End If

    End Sub

#End Region

End Class
