Public Class ErrorReport
    Inherits System.Windows.Forms.Form
    Private ex As System.Exception

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef Except As System.Exception)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.ex = Except

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
    Friend WithEvents txtErrorMsg As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtErrorSource As System.Windows.Forms.TextBox
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents txtStackTrace As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtErrorMsg = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtErrorSource = New System.Windows.Forms.TextBox
        Me.btnOK = New System.Windows.Forms.Button
        Me.txtStackTrace = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtErrorMsg
        '
        Me.txtErrorMsg.Location = New System.Drawing.Point(8, 80)
        Me.txtErrorMsg.Multiline = True
        Me.txtErrorMsg.Name = "txtErrorMsg"
        Me.txtErrorMsg.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtErrorMsg.Size = New System.Drawing.Size(480, 88)
        Me.txtErrorMsg.TabIndex = 0
        Me.txtErrorMsg.Text = "TextBox1"
        Me.txtErrorMsg.WordWrap = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(464, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "An error has occurred in the MUSTER application.  Please note the following:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(264, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Error Message :"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 184)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(296, 16)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Error Source Information"
        '
        'txtErrorSource
        '
        Me.txtErrorSource.Location = New System.Drawing.Point(8, 208)
        Me.txtErrorSource.Multiline = True
        Me.txtErrorSource.Name = "txtErrorSource"
        Me.txtErrorSource.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtErrorSource.Size = New System.Drawing.Size(480, 88)
        Me.txtErrorSource.TabIndex = 4
        Me.txtErrorSource.Text = ""
        Me.txtErrorSource.WordWrap = False
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(208, 456)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(72, 24)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "&OK"
        '
        'txtStackTrace
        '
        Me.txtStackTrace.Location = New System.Drawing.Point(8, 341)
        Me.txtStackTrace.Multiline = True
        Me.txtStackTrace.Name = "txtStackTrace"
        Me.txtStackTrace.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtStackTrace.Size = New System.Drawing.Size(480, 88)
        Me.txtStackTrace.TabIndex = 7
        Me.txtStackTrace.Text = ""
        Me.txtStackTrace.WordWrap = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 317)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(296, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Error Stack Information"
        '
        'ErrorReport
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(496, 494)
        Me.Controls.Add(Me.txtStackTrace)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtErrorSource)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtErrorMsg)
        Me.Name = "ErrorReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ErrorReport"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ErrorReport_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        Try
            txtErrorMsg.Text = ex.Message

            If Not IsNothing(ex.InnerException) Then

                Dim Inner As Exception = ex

                While Not IsNothing(Inner.InnerException)
                    Inner = Inner.InnerException
                End While

                txtErrorSource.Text = Inner.Source
                txtStackTrace.Text = Inner.StackTrace
            Else
                txtErrorSource.Text = ex.Source
                txtErrorSource.Text = ex.StackTrace
            End If
        Catch ex As Exception
            txtErrorMsg.Text = "Error occured while showing the Error Report : " + ex.Message
            txtErrorSource.Text = ex.Source
            txtErrorSource.Text = ex.StackTrace
        End Try

        'txtErrorSource.Text = IIf(ex.Source = String.Empty, ex.InnerException.Source, ex.Source)
        ' txtStackTrace.Text = IIf(ex.StackTrace = String.Empty, ex.InnerException.StackTrace, ex.StackTrace)

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Me.Dispose()

    End Sub
End Class
