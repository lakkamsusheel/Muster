Public Class SalutationSelection
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents LblGenderbase As System.Windows.Forms.Label
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents LblProfessionBase As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.LblGenderbase = New System.Windows.Forms.Label
        Me.LblProfessionBase = New System.Windows.Forms.Label
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.Button9 = New System.Windows.Forms.Button
        Me.Button10 = New System.Windows.Forms.Button
        Me.Button11 = New System.Windows.Forms.Button
        Me.Button12 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 56)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Mr."
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(96, 56)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Mrs."
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(184, 56)
        Me.Button3.Name = "Button3"
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "MS."
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(272, 56)
        Me.Button4.Name = "Button4"
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "Sna."
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(360, 56)
        Me.Button5.Name = "Button5"
        Me.Button5.TabIndex = 4
        Me.Button5.Text = "Snr."
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(448, 56)
        Me.Button6.Name = "Button6"
        Me.Button6.TabIndex = 5
        Me.Button6.Text = "Sr."
        '
        'LblGenderbase
        '
        Me.LblGenderbase.Location = New System.Drawing.Point(8, 32)
        Me.LblGenderbase.Name = "LblGenderbase"
        Me.LblGenderbase.Size = New System.Drawing.Size(128, 23)
        Me.LblGenderbase.TabIndex = 6
        Me.LblGenderbase.Text = "Salutation By Gender"
        '
        'LblProfessionBase
        '
        Me.LblProfessionBase.Location = New System.Drawing.Point(8, 104)
        Me.LblProfessionBase.Name = "LblProfessionBase"
        Me.LblProfessionBase.Size = New System.Drawing.Size(128, 23)
        Me.LblProfessionBase.TabIndex = 13
        Me.LblProfessionBase.Text = "Salutation By Status"
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(448, 128)
        Me.Button7.Name = "Button7"
        Me.Button7.TabIndex = 12
        Me.Button7.Text = "Officer"
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(360, 128)
        Me.Button8.Name = "Button8"
        Me.Button8.TabIndex = 11
        Me.Button8.Text = "Cpt."
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(272, 128)
        Me.Button9.Name = "Button9"
        Me.Button9.TabIndex = 10
        Me.Button9.Text = "Sir"
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(176, 128)
        Me.Button10.Name = "Button10"
        Me.Button10.TabIndex = 9
        Me.Button10.Text = "Mahd"
        '
        'Button11
        '
        Me.Button11.Location = New System.Drawing.Point(96, 128)
        Me.Button11.Name = "Button11"
        Me.Button11.TabIndex = 8
        Me.Button11.Text = "Rev."
        '
        'Button12
        '
        Me.Button12.Location = New System.Drawing.Point(8, 128)
        Me.Button12.Name = "Button12"
        Me.Button12.TabIndex = 7
        Me.Button12.Text = "Dr."
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(496, 24)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "Please Select the Following "
        '
        'SalutationSelection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(528, 174)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.LblProfessionBase)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.LblGenderbase)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SalutationSelection"
        Me.Text = "Salutation Respect Window"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "public properties"

    Public Title As String = "Mr."

    Public Property Header() As String

        Get
            Return Me.Text
        End Get

        Set(ByVal Value As String)
            Text = Value
        End Set

    End Property


#End Region




#Region "Popup Function"

    Public Shared Function SalutationBox(ByVal header As String) As String

        Dim frm As New SalutationSelection


        frm.Header = IIf(header <> String.Empty, header, frm.Header)

        If frm.ShowDialog = frm.DialogResult.OK Then

            Return frm.Title

        Else
            Return String.Empty
        End If

    End Function

#End Region

#Region "Form Events"

    Private Sub SalutationSelection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        AddHandler Button1.Click, AddressOf ButtonClick
        AddHandler Button2.Click, AddressOf ButtonClick
        AddHandler Button3.Click, AddressOf ButtonClick
        AddHandler Button4.Click, AddressOf ButtonClick
        AddHandler Button5.Click, AddressOf ButtonClick
        AddHandler Button6.Click, AddressOf ButtonClick
        AddHandler Button7.Click, AddressOf ButtonClick
        AddHandler Button8.Click, AddressOf ButtonClick
        AddHandler Button9.Click, AddressOf ButtonClick
        AddHandler Button10.Click, AddressOf ButtonClick
        AddHandler Button11.Click, AddressOf ButtonClick
        AddHandler Button12.Click, AddressOf ButtonClick


    End Sub

    Sub ButtonClick(ByVal sender As Object, ByVal e As EventArgs)

        With DirectCast(sender, Button)

            Title = .Text.Trim

            DialogResult = DialogResult.OK

            Close()

        End With

    End Sub

#End Region



End Class
