Public Class pictureViewer
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Event pictureRemovedEvent(ByVal sender As Object, ByVal e As EventArgs)

    Public Sub New(ByVal filePath As IO.FileInfo, ByVal img As Image)

        Me.New()

        pctPicture.Image = img
        With filePath
            lblPicture.Text = String.Format("Filename: {1}{0}Last Modified: {2}{0}Created By: {3}", vbCrLf, .Name, .LastWriteTime.ToLongDateString, String.Empty)
        End With

        FileRemovalButton1.filePath = filePath.FullName

    End Sub

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
    Public WithEvents pctPicture As System.Windows.Forms.PictureBox
    Public WithEvents lblPicture As System.Windows.Forms.Label
    Friend WithEvents FileRemovalButton1 As MUSTER.FileRemovalButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pctPicture = New System.Windows.Forms.PictureBox
        Me.lblPicture = New System.Windows.Forms.Label
        Me.FileRemovalButton1 = New MUSTER.FileRemovalButton
        Me.SuspendLayout()
        '
        'pctPicture
        '
        Me.pctPicture.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pctPicture.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pctPicture.Location = New System.Drawing.Point(8, 8)
        Me.pctPicture.Name = "pctPicture"
        Me.pctPicture.Size = New System.Drawing.Size(640, 588)
        Me.pctPicture.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pctPicture.TabIndex = 0
        Me.pctPicture.TabStop = False
        '
        'lblPicture
        '
        Me.lblPicture.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPicture.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.lblPicture.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPicture.Location = New System.Drawing.Point(648, 32)
        Me.lblPicture.Name = "lblPicture"
        Me.lblPicture.Size = New System.Drawing.Size(304, 560)
        Me.lblPicture.TabIndex = 1
        '
        'FileRemovalButton1
        '
        Me.FileRemovalButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FileRemovalButton1.Location = New System.Drawing.Point(648, 8)
        Me.FileRemovalButton1.Name = "FileRemovalButton1"
        Me.FileRemovalButton1.Size = New System.Drawing.Size(192, 23)
        Me.FileRemovalButton1.TabIndex = 2
        Me.FileRemovalButton1.Text = "Remove Picture"
        '
        'pictureViewer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(952, 597)
        Me.Controls.Add(Me.FileRemovalButton1)
        Me.Controls.Add(Me.lblPicture)
        Me.Controls.Add(Me.pctPicture)
        Me.MaximumSize = New System.Drawing.Size(1280, 1024)
        Me.MinimumSize = New System.Drawing.Size(600, 400)
        Me.Name = "pictureViewer"
        Me.Text = "Image Details"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FileRemovalButton1_Fileremoved(ByVal sender As Object, ByVal e As System.EventArgs) Handles FileRemovalButton1.Fileremoved
        RaiseEvent pictureRemovedEvent(Me, New EventArgs)
        Close()
    End Sub
End Class
