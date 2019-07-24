Public Class WorkShopDate
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
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dtWorkshopdate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.dtWorkshopdate = New System.Windows.Forms.DateTimePicker
        Me.btnOK = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'dtWorkshopdate
        '
        Me.dtWorkshopdate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtWorkshopdate.Location = New System.Drawing.Point(24, 35)
        Me.dtWorkshopdate.Name = "dtWorkshopdate"
        Me.dtWorkshopdate.ShowCheckBox = True
        Me.dtWorkshopdate.Size = New System.Drawing.Size(128, 20)
        Me.dtWorkshopdate.TabIndex = 0
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(16, 72)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(56, 23)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "OK"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(152, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Enter Workshop Date:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(104, 72)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(56, 23)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'WorkShopDate
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(176, 101)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.dtWorkshopdate)
        Me.Controls.Add(Me.btnCancel)
        Me.Name = "WorkShopDate"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Work Shop Date"
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "User defined variable"
    Dim dt As Date = CDate("01/01/0001")
    Private bolAllowClose As Boolean
#End Region
    Public ReadOnly Property WorkshopDate() As Date
        Get
            Return dt
        End Get
    End Property
    Public Property DateLabel() As String
        Get
            Return Label1.Text
        End Get
        Set(ByVal Value As String)
            Label1.Text = Value
        End Set
    End Property
    Public Property frmText() As String
        Get
            Return Me.Text
        End Get
        Set(ByVal Value As String)
            Me.Text = Value
        End Set
    End Property
    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        bolAllowClose = True
        Me.Close()
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        bolAllowClose = True
        dt = CDate("01/01/0001")
        Me.Close()
    End Sub

    Private Sub WorkShopDate_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Text = "Work Shop Date" Then
            btnCancel.Visible = False
        End If
        bolAllowClose = False
        UIUtilsGen.ToggleDateFormat(dtWorkshopdate)
        UIUtilsGen.SetDatePickerValue(dtWorkshopdate, dt)
    End Sub

    Private Sub dtWorkshopdate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtWorkshopdate.ValueChanged
        UIUtilsGen.ToggleDateFormat(dtWorkshopdate)
        dt = UIUtilsGen.GetDatePickerValue(dtWorkshopdate)
        dt = dt.Date
    End Sub

    Private Sub WorkShopDate_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Not bolAllowClose Then
            If Me.Text = "Work Shop Date" Then
                MsgBox("Work Shop Date is required")
                e.Cancel = True
            ElseIf Not MsgBox("Do you want to cancel entering the Date?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
                e.Cancel = True
            End If
        End If
    End Sub
End Class
