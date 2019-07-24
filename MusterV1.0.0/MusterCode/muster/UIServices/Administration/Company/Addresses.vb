Public Class Addresses
    Inherits System.Windows.Forms.Form
#Region "Public events"
    Public Event evtCompLicAssocChanged(ByVal AddressID As Integer)
    Public Event evtProEngAddChanged(ByVal AddressID As Integer)
    Public Event evtProGeoAddChanged(ByVal AddressID As Integer)
#End Region
#Region "User defined Variables"
    Dim ComAdd As MUSTER.BusinessLogic.pComAddress
    Dim CompanyLicensee As MUSTER.BusinessLogic.pCompanyLicensee
    Dim pCompany As MUSTER.BusinessLogic.pCompany
    Dim nAddressID As Integer
    Dim strFrom As String = ""
    Dim nAssocID As Integer = 0
    Dim nCompanyID As Integer
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef CompanyAddresses As MUSTER.BusinessLogic.pComAddress, ByVal strForm As String, ByVal AssocID As String, Optional ByRef oCompanyLicensee As MUSTER.BusinessLogic.pCompanyLicensee = Nothing, Optional ByRef company As MUSTER.BusinessLogic.pCompany = Nothing, Optional ByVal CompanyID As Integer = 0)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        ComAdd = CompanyAddresses
        strFrom = strForm
        nCompanyID = CompanyID
        If Not oCompanyLicensee Is Nothing Then
            CompanyLicensee = oCompanyLicensee
        End If
        If Not company Is Nothing Then
            pCompany = company
        End If
        nAssocID = AssocID
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
    Friend WithEvents ugAddresses As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ugAddresses = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.ugAddresses, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ugAddresses
        '
        Me.ugAddresses.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAddresses.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAddresses.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAddresses.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAddresses.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugAddresses.Location = New System.Drawing.Point(0, 0)
        Me.ugAddresses.Name = "ugAddresses"
        Me.ugAddresses.Size = New System.Drawing.Size(712, 230)
        Me.ugAddresses.TabIndex = 0
        '
        'Addresses
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(712, 230)
        Me.Controls.Add(Me.ugAddresses)
        Me.Name = "Addresses"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Addresses"
        CType(Me.ugAddresses, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Addresses_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim xComAddressInfo As MUSTER.Info.ComAddressInfo
            'ComAdd.GetAddressAll(0, nCompanyID, 0, 0, False)
            'ugAddresses.DataSource = ComAdd.AddressTable.DefaultView
            ugAddresses.DataSource = ComAdd.GetCompanyAddress(nCompanyID) 'ComAdd.AddressTable.DefaultView
            ugAddresses.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            ugAddresses.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_ONE").Width = 100
            ugAddresses.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_TWO").Width = 100
            ugAddresses.DisplayLayout.Bands(0).Columns("CITY").Width = 90
            ugAddresses.DisplayLayout.Bands(0).Columns("STATE").Width = 50
            ugAddresses.DisplayLayout.Bands(0).Columns("ZIP").Width = 70
            ugAddresses.DisplayLayout.Bands(0).Columns("PHONE_NUMBER_ONE").Width = 100
            ugAddresses.DisplayLayout.Bands(0).Columns("EXT_ONE").Width = 60
            ugAddresses.DisplayLayout.Bands(0).Columns("PHONE_NUMBER_TWO").Width = 100
            ugAddresses.DisplayLayout.Bands(0).Columns("EXT_TWO").Width = 60
            'ugAddresses.DisplayLayout.Bands(0).Columns("ADDRESS_ID").Hidden = True
            ugAddresses.DisplayLayout.Bands(0).Columns("COM_ADDRESS_ID").Hidden = True
            ugAddresses.DisplayLayout.Bands(0).Columns("COMPANY_ID").Hidden = True
            ugAddresses.DisplayLayout.Bands(0).Columns("LICENSEE_ID").Hidden = True
            ugAddresses.DisplayLayout.Bands(0).Columns("PROVIDER_ID").Hidden = True


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Caption
            ' Address 1, Address 2, City, State, Zip, Phone 1, Ext1, Phone 2, Ext2
            ugAddresses.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_ONE").Header.Caption = "Address 1"
            ugAddresses.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_TWO").Header.Caption = "Address 2"
            ugAddresses.DisplayLayout.Bands(0).Columns("CITY").Header.Caption = "City"
            ugAddresses.DisplayLayout.Bands(0).Columns("STATE").Header.Caption = "State"
            ugAddresses.DisplayLayout.Bands(0).Columns("ZIP").Header.Caption = "Zip"
            ugAddresses.DisplayLayout.Bands(0).Columns("PHONE_NUMBER_ONE").Header.Caption = "Phone 1"
            ugAddresses.DisplayLayout.Bands(0).Columns("EXT_ONE").Header.Caption = "Ext1"
            ugAddresses.DisplayLayout.Bands(0).Columns("PHONE_NUMBER_TWO").Header.Caption = "Phone 2"
            ugAddresses.DisplayLayout.Bands(0).Columns("EXT_TWO").Header.Caption = "Ext2"

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugAddresses_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAddresses.DoubleClick
        Try
            If (strFrom = "Licensee") OrElse (strFrom = "Manager") Then
                If nAssocID > 0 Or nAssocID < -100 Then
                    RaiseEvent evtCompLicAssocChanged(ugAddresses.ActiveRow.Cells("COM_ADDRESS_ID").Value)
                Else
                    CompanyLicensee.ComLicAddressID = ugAddresses.ActiveRow.Cells("COM_ADDRESS_ID").Value
                End If
            End If
            If strFrom = "CompanyProEngineer" Then
                RaiseEvent evtProEngAddChanged(ugAddresses.ActiveRow.Cells("COM_ADDRESS_ID").Value)
            End If
            If strFrom = "CompanyProGeologist" Then
                RaiseEvent evtProGeoAddChanged(ugAddresses.ActiveRow.Cells("COM_ADDRESS_ID").Value)
            End If
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
