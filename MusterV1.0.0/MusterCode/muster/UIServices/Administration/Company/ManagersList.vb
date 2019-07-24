Public Class ManagersList
    Inherits System.Windows.Forms.Form
#Region "Public Events"
    Public Event CompanyManagerAssociation(ByVal ManagerID As Integer)
#End Region
#Region "User defined variables"
    Private pManager As MUSTER.BusinessLogic.pLicensee
    Private ManagerInfo As MUSTER.Info.LicenseeInfo
    Private pCompany As MUSTER.BusinessLogic.pCompany
    Dim nManagerID As Integer = 0
    Friend WithEvents objLicen As Managers
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef oCompany As MUSTER.BusinessLogic.pCompany, ByRef oManager As MUSTER.BusinessLogic.pLicensee)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        pCompany = oCompany
        pManager = oManager
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
    Friend WithEvents ugManagers As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ugManagers = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.ugManagers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ugManagers
        '
        Me.ugManagers.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugManagers.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugManagers.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugManagers.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugManagers.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugManagers.Location = New System.Drawing.Point(0, 0)
        Me.ugManagers.Name = "ugManagers"
        Me.ugManagers.Size = New System.Drawing.Size(744, 270)
        Me.ugManagers.TabIndex = 0
        '
        'ManagersList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(744, 270)
        Me.Controls.Add(Me.ugManagers)
        Me.Name = "ManagersList"
        Me.Text = "ManagersList"
        CType(Me.ugManagers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub ManagersList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dsData As DataSet
        Try
            If pCompany Is Nothing Then
                dsData = pManager.GetManagerList()
            Else
                dsData = pManager.GetManagerList(pCompany.ID)
            End If

            ugManagers.DataSource = dsData.Tables(0).DefaultView
            ugManagers.DisplayLayout.Bands(0).Columns("Manager_ID").Hidden = True

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugManagers_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugManagers.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not (ugManagers.ActiveRow Is Nothing) Then
                nManagerID = ugManagers.ActiveRow.Cells("Manager_ID").Value
            Else
                Exit Sub
            End If
            'ManagerInfo = New MUSTER.Info.ManagerInfo(ugManagers.ActiveRow.Cells("Manager_ID").Value, _
            '                                            ugManagers.ActiveRow.Cells("TITLE").Value, _
            '                                            ugManagers.ActiveRow.Cells("FIRST_NAME").Value, _
            '                                            ugManagers.ActiveRow.Cells("MIDDLE_NAME").Value, _
            '                                            ugManagers.ActiveRow.Cells("LAST_NAME").Value, _
            '                                            ugManagers.ActiveRow.Cells("SUFFIX").Value, _
            '                                            ugManagers.ActiveRow.Cells("LICENSE_NUMBER_PREFIX").Value, _
            '                                            ugManagers.ActiveRow.Cells("LICENSE_NUMBER").Value, _
            '                                            ugManagers.ActiveRow.Cells("EMAIL_ADDRESS").Value, _
            '                                            ugManagers.ActiveRow.Cells("ASSOCATED_COMPANY_ID").Value, _
            '                                            ugManagers.ActiveRow.Cells("HIRE_STATUS").Value, _
            '                                            ugManagers.ActiveRow.Cells("EMPLOYEE_LETTER").Value, _
            '                                            ugManagers.ActiveRow.Cells("STATUS").Value, _
            '                                            ugManagers.ActiveRow.Cells("OVERRIDE_EXPIRE").Value, _
            '                                            ugManagers.ActiveRow.Cells("CERT_TYPE").Value, _
            '                                            ugManagers.ActiveRow.Cells("APP_RECVD_DATE").Value, _
            '                                            ugManagers.ActiveRow.Cells("ORGIN_ISSUED_DATE").Value, _
            '                                            ugManagers.ActiveRow.Cells("ISSUED_DATE").Value, _
            '                                            ugManagers.ActiveRow.Cells("LICENSE_EXPIRE_DATE").Value, _
            '                                            ugManagers.ActiveRow.Cells("EXCEPT_GRANT_DATE").Value, _
            '                                            ugManagers.ActiveRow.Cells("CREATED_BY").Value, _
            '                                            ugManagers.ActiveRow.Cells("DATE_CREATED").Value, _
            '                                            IIf(ugManagers.ActiveRow.Cells("LAST_EDITED_BY").Value Is System.DBNull.Value, String.Empty, ugManagers.ActiveRow.Cells("LAST_EDITED_BY").Value), _
            '                                            IIf(ugManagers.ActiveRow.Cells("DATE_LAST_EDITED").Value Is System.DBNull.Value, CDate("01/01/0001"), ugManagers.ActiveRow.Cells("DATE_LAST_EDITED").Value), _
            '                                            ugManagers.ActiveRow.Cells("DELETED").Value)
            'pManager.Add(ManagerInfo)
            Me.Close()
            RaiseEvent CompanyManagerAssociation(nManagerID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

End Class
