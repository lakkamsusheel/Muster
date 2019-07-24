Public Class LicenseesList
    Inherits System.Windows.Forms.Form
#Region "Public Events"
    Public Event CompanyLicenseeAssociation(ByVal LicenseeID As Integer)
#End Region
#Region "User defined variables"
    Private pLicensee As MUSTER.BusinessLogic.pLicensee
    Private LicenseeInfo As MUSTER.Info.LicenseeInfo
    Private pCompany As MUSTER.BusinessLogic.pCompany
    Dim nLicenseeID As Integer = 0
    Friend WithEvents objLicen As Licensees
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef oCompany As MUSTER.BusinessLogic.pCompany, ByRef oLicensee As MUSTER.BusinessLogic.pLicensee)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        pCompany = oCompany
        pLicensee = oLicensee
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
        Me.ugLicensees.Size = New System.Drawing.Size(744, 270)
        Me.ugLicensees.TabIndex = 0
        '
        'LicenseesList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(744, 270)
        Me.Controls.Add(Me.ugLicensees)
        Me.Name = "LicenseesList"
        Me.Text = "LicenseesList"
        CType(Me.ugLicensees, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub LicenseesList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dsData As DataSet
        Try
            If pCompany Is Nothing Then
                dsData = pLicensee.GetLicenseeList()
            Else
                dsData = pLicensee.GetLicenseeList(pCompany.ID)
            End If

            ugLicensees.DataSource = dsData.Tables(0).DefaultView
            ugLicensees.DisplayLayout.Bands(0).Columns("Licensee_ID").Hidden = True

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugLicensees_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugLicensees.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not (ugLicensees.ActiveRow Is Nothing) Then
                nLicenseeID = ugLicensees.ActiveRow.Cells("Licensee_ID").Value
            Else
                Exit Sub
            End If
            'LicenseeInfo = New MUSTER.Info.LicenseeInfo(ugLicensees.ActiveRow.Cells("LICENSEE_ID").Value, _
            '                                            ugLicensees.ActiveRow.Cells("TITLE").Value, _
            '                                            ugLicensees.ActiveRow.Cells("FIRST_NAME").Value, _
            '                                            ugLicensees.ActiveRow.Cells("MIDDLE_NAME").Value, _
            '                                            ugLicensees.ActiveRow.Cells("LAST_NAME").Value, _
            '                                            ugLicensees.ActiveRow.Cells("SUFFIX").Value, _
            '                                            ugLicensees.ActiveRow.Cells("LICENSE_NUMBER_PREFIX").Value, _
            '                                            ugLicensees.ActiveRow.Cells("LICENSE_NUMBER").Value, _
            '                                            ugLicensees.ActiveRow.Cells("EMAIL_ADDRESS").Value, _
            '                                            ugLicensees.ActiveRow.Cells("ASSOCATED_COMPANY_ID").Value, _
            '                                            ugLicensees.ActiveRow.Cells("HIRE_STATUS").Value, _
            '                                            ugLicensees.ActiveRow.Cells("EMPLOYEE_LETTER").Value, _
            '                                            ugLicensees.ActiveRow.Cells("STATUS").Value, _
            '                                            ugLicensees.ActiveRow.Cells("OVERRIDE_EXPIRE").Value, _
            '                                            ugLicensees.ActiveRow.Cells("CERT_TYPE").Value, _
            '                                            ugLicensees.ActiveRow.Cells("APP_RECVD_DATE").Value, _
            '                                            ugLicensees.ActiveRow.Cells("ORGIN_ISSUED_DATE").Value, _
            '                                            ugLicensees.ActiveRow.Cells("ISSUED_DATE").Value, _
            '                                            ugLicensees.ActiveRow.Cells("LICENSE_EXPIRE_DATE").Value, _
            '                                            ugLicensees.ActiveRow.Cells("EXCEPT_GRANT_DATE").Value, _
            '                                            ugLicensees.ActiveRow.Cells("CREATED_BY").Value, _
            '                                            ugLicensees.ActiveRow.Cells("DATE_CREATED").Value, _
            '                                            IIf(ugLicensees.ActiveRow.Cells("LAST_EDITED_BY").Value Is System.DBNull.Value, String.Empty, ugLicensees.ActiveRow.Cells("LAST_EDITED_BY").Value), _
            '                                            IIf(ugLicensees.ActiveRow.Cells("DATE_LAST_EDITED").Value Is System.DBNull.Value, CDate("01/01/0001"), ugLicensees.ActiveRow.Cells("DATE_LAST_EDITED").Value), _
            '                                            ugLicensees.ActiveRow.Cells("DELETED").Value)
            'pLicensee.Add(LicenseeInfo)
            Me.Close()
            RaiseEvent CompanyLicenseeAssociation(nLicenseeID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

End Class
