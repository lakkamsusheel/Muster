Public Class ProgressScreen
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
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(32, 40)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(416, 23)
        Me.ProgressBar1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.Location = New System.Drawing.Point(32, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(416, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(360, 80)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Minmize Me"
        '
        'ProgressScreen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 110)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProgressBar1)
        Me.MaximizeBox = False
        Me.Name = "ProgressScreen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Progress Screen"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "private members"

    Private _container As MusterContainer
    Private _maxValue As Long
    Private _value As Long
    Private _title As String
    Private _msg As String
    Private _args() As String

    Private WithEvents _UnzipTimer As System.Timers.Timer

    Public Event ReturningMessage(ByVal msg As String)


#End Region

#Region "Construct"

    Sub New(ByVal container As MusterContainer)

        'Visible = False

        _container = container

        RemoveHandler _container.FireProgressMessage, AddressOf receiveMessage
        AddHandler _container.FireProgressMessage, AddressOf receiveMessage

        RemoveHandler _container.SizeChanged, AddressOf MatchSize
        AddHandler _container.SizeChanged, AddressOf MatchSize

        RemoveHandler _container.StartProgressScreen, AddressOf ShowMe
        AddHandler _container.StartProgressScreen, AddressOf ShowMe

        RemoveHandler _container.FireCloseProgressScreen, AddressOf CloseByContainer
        AddHandler _container.FireCloseProgressScreen, AddressOf CloseByContainer

    End Sub

#End Region


#Region "methods"

    Private Sub PushValues(ByVal percent As Integer, ByVal message As String)



        ProgressBar1.Value = percent
        Me.Label1.Text = message

        Me.ProgressBar1.Refresh()

        Me.Label1.Refresh()

        Me.Refresh()


    End Sub

    Private Sub ShowMe(ByVal title As String, ByVal maxvalue As Long, ByVal val As Long, ByVal msg As String, ByVal ParamArray codeArgs() As String)

        Width = 488
        Height = 144

        Me.TopMost = True

        _title = title
        _maxValue = maxvalue
        _value = val
        _msg = msg

        Show()

        SetForm()

        Update()

        processCode(True, codeArgs)

    End Sub

    Sub SetForm()

        ProgressBar1.Maximum = _maxValue
        ProgressBar1.Value = _value

        Label1.Text = _msg

        ControlBox = False

        Text = _title

        _container.Enabled = False

    End Sub

    Sub ShowForm(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

        If ProgressBar1 Is Nothing Then

            ProgressBar1 = New ProgressBar
            ProgressBar1.Location = New Point(32, 40)
            ProgressBar1.Size = New Size(416, 23)

            Me.Controls.Add(ProgressBar1)

        End If

        If Label1 Is Nothing Then

            Label1 = New Label
            Label1.BackColor = Color.WhiteSmoke
            Label1.BorderStyle = BorderStyle.Fixed3D
            Label1.Location = New Point(32, 8)
            Label1.Size = New Size(416, 23)

            Me.Controls.Add(Label1)
        End If

        If Button1 Is Nothing Then
            Button1 = New Button
            Button1.Text = "Minimize Me"
            Button1.Location = New Point(360, 80)
            Button1.Size = New Size(88, 23)

            Me.Controls.Add(Button1)
        End If

    End Sub

    Sub CloseMe()


        If Not _container Is Nothing Then
            _container.Enabled = True
        End If

        Me.Hide()

    End Sub

    Private Sub processCode(ByVal SetArgs As Boolean, ByVal ParamArray code() As String)

        If code.GetUpperBound(0) >= 0 Then

            Dim firstCode As String = code(0)

            If SetArgs Then
                _args = code
            End If

            If Not firstCode Is Nothing AndAlso firstCode.Length > 0 Then

                Select Case firstCode
                    Case "PrepareUnzip"
                        StartUnzipTimer()
                        RaiseEvent ReturningMessage("ReadyForUnzip")
                    Case "CompleteUnzip"
                        _UnzipTimer.Stop()
                        _UnzipTimer = Nothing
                        RaiseEvent ReturningMessage("UnzipCompletedAcknowledged")
                    Case "RestoreCompleted"
                        RaiseEvent ReturningMessage("RestoreCompletedAcknowledged")



                End Select

            End If
        End If


    End Sub

    Sub StartUnzipTimer()

        _UnzipTimer = New System.Timers.Timer(1000)

        _UnzipTimer.AutoReset = False

        _UnzipTimer.Start()


    End Sub

#End Region

#Region "Action Events"

    Private Sub UpdateZipProgress(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles _UnzipTimer.Elapsed

        With DirectCast(sender, System.Timers.Timer)

            Dim size As Integer = 0
            Dim maxSize As Int64 = 0

            If IO.File.Exists(_args(2)) Then
                size = (FileSystem.FileLen(_args(2)) / 100000)
            End If

            maxSize = Convert.ToInt64(_args(1)) / 100000

            If size > 0 Then
                PushValues(size, String.Format("Unzipping backup file:  {0:#.##}%", (size / maxSize) * 100))
            Else
                PushValues(0, String.Format("Unzipping backup file: 0 %"))
            End If

            _UnzipTimer.Start()

        End With

    End Sub

#End Region

#Region "Events"


    Private Sub CloseByContainer()
        CloseMe()
    End Sub
    Private Sub receiveMessage(ByVal msg As String, ByVal num As Long, Optional ByVal code As String = "")

        PushValues(num, msg)

        processCode(False, code)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        WindowState = FormWindowState.Minimized

    End Sub

    Private Sub MatchSize(ByVal sender As Object, ByVal e As System.EventArgs)

        If Not _container Is Nothing AndAlso Not _container.Enabled Then

            If _container.WindowState <> FormWindowState.Minimized AndAlso Me.WindowState = FormWindowState.Minimized Then

                Me.WindowState = FormWindowState.Normal

            ElseIf _container.WindowState = FormWindowState.Minimized AndAlso Me.WindowState <> FormWindowState.Minimized Then

                Me.WindowState = FormWindowState.Minimized

            End If

        End If

    End Sub

    Private Sub ProgressScreen_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged

        If Not _container Is Nothing AndAlso Not _container.Enabled Then

            If Me.WindowState <> FormWindowState.Minimized AndAlso _container.WindowState = FormWindowState.Minimized Then

                _container.WindowState = FormWindowState.Maximized

            ElseIf Me.WindowState = FormWindowState.Minimized AndAlso _container.WindowState <> FormWindowState.Minimized Then

                _container.WindowState = FormWindowState.Minimized

            End If

            _container.Refresh()
        End If


    End Sub
#End Region



End Class
