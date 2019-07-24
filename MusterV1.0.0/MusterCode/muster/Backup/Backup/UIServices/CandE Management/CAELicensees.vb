Public Class CAELicensees
    Inherits System.Windows.Forms.Form
#Region "User defined variables"
    Private pLCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef LCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        pLCE = LCE
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
    Friend WithEvents ugLicensees As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ugLicensees = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.ugLicensees, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ugLicensees
        '
        Me.ugLicensees.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugLicensees.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugLicensees.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugLicensees.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugLicensees.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugLicensees.Location = New System.Drawing.Point(0, 0)
        Me.ugLicensees.Name = "ugLicensees"
        Me.ugLicensees.Size = New System.Drawing.Size(792, 270)
        Me.ugLicensees.TabIndex = 1
        '
        'CAELicensees
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(792, 270)
        Me.Controls.Add(Me.ugLicensees)
        Me.Name = "CAELicensees"
        Me.Text = "CAELicensees"
        CType(Me.ugLicensees, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CAELicensees_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ugLicensees.DataSource = pLCE.getLCELicensees()
            ugLicensees.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            ugLicensees.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_ONE").Hidden = True
            ugLicensees.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_TWO").Hidden = True
            ugLicensees.DisplayLayout.Bands(0).Columns("CITY").Hidden = True
            ugLicensees.DisplayLayout.Bands(0).Columns("STATE").Hidden = True
            ugLicensees.DisplayLayout.Bands(0).Columns("ZIP").Hidden = True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugLicensees_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugLicensees.DoubleClick
        Try
            If Not (ugLicensees.ActiveRow Is Nothing) Then
                pLCE.LicenseeID = ugLicensees.ActiveRow.Cells("LICENSEE_ID").Value
                pLCE.LicenseeName = ugLicensees.ActiveRow.Cells("LICENSEE_NAME").Value
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
