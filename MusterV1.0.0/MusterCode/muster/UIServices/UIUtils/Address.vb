Imports System.Configuration

Public Class Address
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Public WithEvents pAddress As MUSTER.BusinessLogic.pAddress
    Private bolShowCounty As Boolean = True
    Private bolShowFIPS As Boolean = True
    Private bolLoading As Boolean = False
    Private nownerAddressID, nEntityType, nModuleID, nEntityID As Integer
    Dim returnVal As String = String.Empty
#End Region
#Region "User Defined Events"
    Public Event NewAddressID(ByVal MyAddressID As Int32)
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(ByVal entityType As Integer, ByVal entity As Integer, Optional ByRef pAddr As MUSTER.BusinessLogic.pAddress = Nothing, Optional ByVal callingEntity As String = "", Optional ByVal ownerAddressID As Integer = 0, Optional ByVal moduleID As Integer = 0, Optional ByVal isReadOnly As Boolean = False)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        If pAddr Is Nothing Then
            pAddress = New MUSTER.BusinessLogic.pAddress
        Else
            pAddress = pAddr
        End If

        If pAddress.Zip.Length >= 5 Then
            pAddress.Zip = pAddress.Zip.Substring(0, 5)
        End If

        nownerAddressID = ownerAddressID
        nEntityType = entityType
        nEntityID = entity
        ' if owner addressid is not provided (0), do not show copy from owner checkbox
        chkBoxCopyFromOwner.Visible = nownerAddressID
        If callingEntity <> String.Empty Then
            Me.Text = callingEntity + " " + Me.Text
        End If
        nModuleID = moduleID
        If nModuleID = 0 Then nModuleID = UIUtilsGen.ModuleID.Registration
        LoadAddress()

        If isReadOnly Then
            For Each con As Control In Me.Controls
                con.Enabled = False
            Next
            btnCancel.Enabled = True
            btnMap.Enabled = True

        End If

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
    Friend WithEvents btnClearData As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnAccept As System.Windows.Forms.Button
    Friend WithEvents lblZip As System.Windows.Forms.Label
    Friend WithEvents lblState As System.Windows.Forms.Label
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblCounty As System.Windows.Forms.Label
    Friend WithEvents lblFIPS As System.Windows.Forms.Label
    Friend WithEvents lblAddress2 As System.Windows.Forms.Label
    Friend WithEvents lblAddress1 As System.Windows.Forms.Label
    Friend WithEvents cboCounty As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboZipCode As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboState As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboCity As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboFIPS As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents txtAddress2 As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress1 As System.Windows.Forms.TextBox
    Friend WithEvents chkBoxCopyFromOwner As System.Windows.Forms.CheckBox
    Friend WithEvents TxtPhysicalTown As System.Windows.Forms.TextBox
    Friend WithEvents lblPhysicalTown As System.Windows.Forms.Label
    Friend WithEvents btnMap As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnClearData = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnAccept = New System.Windows.Forms.Button
        Me.lblZip = New System.Windows.Forms.Label
        Me.lblState = New System.Windows.Forms.Label
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblCounty = New System.Windows.Forms.Label
        Me.lblFIPS = New System.Windows.Forms.Label
        Me.lblAddress2 = New System.Windows.Forms.Label
        Me.lblAddress1 = New System.Windows.Forms.Label
        Me.cboCounty = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboZipCode = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboState = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboCity = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboFIPS = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.txtAddress2 = New System.Windows.Forms.TextBox
        Me.txtAddress1 = New System.Windows.Forms.TextBox
        Me.chkBoxCopyFromOwner = New System.Windows.Forms.CheckBox
        Me.TxtPhysicalTown = New System.Windows.Forms.TextBox
        Me.lblPhysicalTown = New System.Windows.Forms.Label
        Me.btnMap = New System.Windows.Forms.Button
        CType(Me.cboCounty, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboZipCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboState, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboCity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboFIPS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnClearData
        '
        Me.btnClearData.Location = New System.Drawing.Point(88, 256)
        Me.btnClearData.Name = "btnClearData"
        Me.btnClearData.Size = New System.Drawing.Size(72, 24)
        Me.btnClearData.TabIndex = 10
        Me.btnClearData.Text = "C&lear"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(248, 256)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(72, 24)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "&Cancel"
        '
        'btnAccept
        '
        Me.btnAccept.Enabled = False
        Me.btnAccept.Location = New System.Drawing.Point(8, 256)
        Me.btnAccept.Name = "btnAccept"
        Me.btnAccept.Size = New System.Drawing.Size(72, 24)
        Me.btnAccept.TabIndex = 4
        Me.btnAccept.Text = "&Accept"
        '
        'lblZip
        '
        Me.lblZip.Location = New System.Drawing.Point(8, 152)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(80, 16)
        Me.lblZip.TabIndex = 0
        Me.lblZip.Text = "Zip Code"
        Me.lblZip.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblState
        '
        Me.lblState.Location = New System.Drawing.Point(8, 128)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(80, 16)
        Me.lblState.TabIndex = 0
        Me.lblState.Text = "State"
        Me.lblState.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(8, 104)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(80, 16)
        Me.lblCity.TabIndex = 0
        Me.lblCity.Text = "City"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCounty
        '
        Me.lblCounty.Location = New System.Drawing.Point(8, 80)
        Me.lblCounty.Name = "lblCounty"
        Me.lblCounty.Size = New System.Drawing.Size(80, 16)
        Me.lblCounty.TabIndex = 0
        Me.lblCounty.Text = "County"
        Me.lblCounty.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblFIPS
        '
        Me.lblFIPS.Location = New System.Drawing.Point(8, 56)
        Me.lblFIPS.Name = "lblFIPS"
        Me.lblFIPS.Size = New System.Drawing.Size(80, 16)
        Me.lblFIPS.TabIndex = 0
        Me.lblFIPS.Text = "FIPS Code"
        Me.lblFIPS.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblAddress2
        '
        Me.lblAddress2.Location = New System.Drawing.Point(8, 32)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(80, 16)
        Me.lblAddress2.TabIndex = 0
        Me.lblAddress2.Text = "Address Line 2"
        Me.lblAddress2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblAddress1
        '
        Me.lblAddress1.Location = New System.Drawing.Point(8, 8)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New System.Drawing.Size(80, 16)
        Me.lblAddress1.TabIndex = 0
        Me.lblAddress1.Text = "Address Line 1"
        Me.lblAddress1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cboCounty
        '
        Me.cboCounty.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboCounty.DisplayMember = ""
        Me.cboCounty.Location = New System.Drawing.Point(96, 80)
        Me.cboCounty.Name = "cboCounty"
        Me.cboCounty.Size = New System.Drawing.Size(200, 21)
        Me.cboCounty.TabIndex = 7
        Me.cboCounty.ValueMember = ""
        '
        'cboZipCode
        '
        Me.cboZipCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboZipCode.DisplayMember = ""
        Me.cboZipCode.Location = New System.Drawing.Point(96, 152)
        Me.cboZipCode.Name = "cboZipCode"
        Me.cboZipCode.Size = New System.Drawing.Size(136, 21)
        Me.cboZipCode.TabIndex = 3
        Me.cboZipCode.ValueMember = ""
        '
        'cboState
        '
        Me.cboState.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboState.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Inset
        Me.cboState.DisplayMember = ""
        Me.cboState.Location = New System.Drawing.Point(96, 128)
        Me.cboState.Name = "cboState"
        Me.cboState.Size = New System.Drawing.Size(72, 21)
        Me.cboState.TabIndex = 5
        Me.cboState.ValueMember = ""
        '
        'cboCity
        '
        Me.cboCity.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboCity.DisplayMember = ""
        Me.cboCity.Location = New System.Drawing.Point(96, 104)
        Me.cboCity.Name = "cboCity"
        Me.cboCity.Size = New System.Drawing.Size(200, 21)
        Me.cboCity.TabIndex = 6
        Me.cboCity.ValueMember = ""
        '
        'cboFIPS
        '
        Me.cboFIPS.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboFIPS.DisplayMember = ""
        Me.cboFIPS.Location = New System.Drawing.Point(96, 56)
        Me.cboFIPS.Name = "cboFIPS"
        Me.cboFIPS.Size = New System.Drawing.Size(96, 21)
        Me.cboFIPS.TabIndex = 8
        Me.cboFIPS.ValueMember = ""
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(96, 32)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(200, 20)
        Me.txtAddress2.TabIndex = 2
        Me.txtAddress2.Text = ""
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New System.Drawing.Point(96, 8)
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(200, 20)
        Me.txtAddress1.TabIndex = 1
        Me.txtAddress1.Text = ""
        '
        'chkBoxCopyFromOwner
        '
        Me.chkBoxCopyFromOwner.Location = New System.Drawing.Point(96, 216)
        Me.chkBoxCopyFromOwner.Name = "chkBoxCopyFromOwner"
        Me.chkBoxCopyFromOwner.Size = New System.Drawing.Size(176, 32)
        Me.chkBoxCopyFromOwner.TabIndex = 0
        Me.chkBoxCopyFromOwner.Text = "Copy Address from Owner"
        '
        'TxtPhysicalTown
        '
        Me.TxtPhysicalTown.Location = New System.Drawing.Point(96, 184)
        Me.TxtPhysicalTown.Name = "TxtPhysicalTown"
        Me.TxtPhysicalTown.Size = New System.Drawing.Size(200, 20)
        Me.TxtPhysicalTown.TabIndex = 11
        Me.TxtPhysicalTown.Text = ""
        '
        'lblPhysicalTown
        '
        Me.lblPhysicalTown.Location = New System.Drawing.Point(8, 184)
        Me.lblPhysicalTown.Name = "lblPhysicalTown"
        Me.lblPhysicalTown.Size = New System.Drawing.Size(80, 16)
        Me.lblPhysicalTown.TabIndex = 12
        Me.lblPhysicalTown.Text = "Actual Town "
        Me.lblPhysicalTown.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnMap
        '
        Me.btnMap.Location = New System.Drawing.Point(168, 256)
        Me.btnMap.Name = "btnMap"
        Me.btnMap.Size = New System.Drawing.Size(72, 24)
        Me.btnMap.TabIndex = 13
        Me.btnMap.Text = "&Map"
        '
        'Address
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(328, 286)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnMap)
        Me.Controls.Add(Me.lblPhysicalTown)
        Me.Controls.Add(Me.TxtPhysicalTown)
        Me.Controls.Add(Me.chkBoxCopyFromOwner)
        Me.Controls.Add(Me.btnClearData)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnAccept)
        Me.Controls.Add(Me.lblZip)
        Me.Controls.Add(Me.lblState)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblCounty)
        Me.Controls.Add(Me.lblFIPS)
        Me.Controls.Add(Me.lblAddress2)
        Me.Controls.Add(Me.lblAddress1)
        Me.Controls.Add(Me.cboCounty)
        Me.Controls.Add(Me.cboZipCode)
        Me.Controls.Add(Me.cboState)
        Me.Controls.Add(Me.cboCity)
        Me.Controls.Add(Me.cboFIPS)
        Me.Controls.Add(Me.txtAddress2)
        Me.Controls.Add(Me.txtAddress1)
        Me.Name = "Address"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Address"
        CType(Me.cboCounty, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboZipCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboState, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboCity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboFIPS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Support Routines"
    Public Sub LoadAddress()
        Try
            If pAddress.AddressId > 0 Then
                pAddress.Retrieve(pAddress.AddressId)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Public Shared Sub EditAddress(ByVal addressform As MUSTER.Address, ByVal ID As Integer, ByRef addresses As BusinessLogic.pAddress, ByVal callingEntity As String, ByVal moduleID As UIUtilsGen.ModuleID, ByRef addressText As TextBox, ByVal entityType As UIUtilsGen.EntityTypes, Optional ByVal isFacilityAddress As Boolean = False)
        Try

            addressform = New Address(entityType, ID, addresses, callingEntity, , moduleID)
            addressform.ShowCounty = False
            addressform.ShowFIPS = False
            addressform.ShowDialog()
            ' update address text
            If addressform.DialogResult = addressform.DialogResult.OK Then

                Dim addressID = addressform.pAddress.AddressId

                addresses = New BusinessLogic.pAddress
                addresses.Retrieve(addressID)

                addressText.Text = UIUtilsGen.FormatAddress(addresses, isFacilityAddress)
                addressText.Tag = addresses.AddressId

            End If

            addressform.Dispose()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub UpdateComboBoxes(ByVal Code As String)
        Try
            Dim strStates, strCities, strCounties, strZips, strFips As String
            If Code = "States" Or Code = String.Empty Then
                strStates = "SELECT DISTINCT STATE FROM tblSYS_ZIPCODES WHERE ZIPID >= 1" + _
                                IIf(pAddress.City.Trim <> "", " AND CITY LIKE '%" + pAddress.City.Trim + "%'", "") + _
                                IIf(bolShowCounty And pAddress.County.Trim <> "", " AND COUNTY LIKE '%" + pAddress.County.Trim + "%'", "") + _
                                IIf(pAddress.Zip.Trim <> "", " AND ZIP LIKE '%" + pAddress.Zip.Trim + "%'", "") + _
                                IIf(bolShowFIPS And pAddress.FIPSCode.Trim <> "", " AND FIPS LIKE '%" + pAddress.FIPSCode.Trim + "%'", "") + _
                                " Union Select ' ' as State " + _
                                " ORDER BY STATE"
                cboState.DataSource = pAddress.GetDataSet(strStates)
                If cboState.Rows.Count = 2 Or cboState.Rows.Count = 1 Then
                    pAddress.State = cboState.Rows.Item(cboState.Rows.Count - 1).Cells(0).Value
                End If
            End If


            If Code = "Cities" Or Code = String.Empty Then
                strCities = "SELECT DISTINCT CITY FROM tblSYS_ZIPCODES WHERE ZIPID >= 1 " + _
                                IIf(pAddress.State.Trim <> "", " and STATE LIKE '%" + pAddress.State.Trim + "%'", "") + _
                                IIf(bolShowCounty And pAddress.County.Trim <> "", " AND COUNTY LIKE '%" + pAddress.County.Trim + "%'", "") + _
                                IIf(pAddress.Zip.Trim <> "", " AND ZIP LIKE '%" + pAddress.Zip.Trim + "%'", "") + _
                                IIf(bolShowFIPS And pAddress.FIPSCode.Trim <> "", " AND FIPS LIKE '%" + pAddress.FIPSCode.Trim + "%'", "") + _
                                 " Union Select ' ' as City " + _
                                " ORDER BY CITY"
                cboCity.DataSource = pAddress.GetDataSet(strCities)
                If cboCity.Rows.Count = 2 Or cboCity.Rows.Count = 1 Then
                    pAddress.City = cboCity.Rows.Item(cboCity.Rows.Count - 1).Cells(0).Value
                End If
            End If


            If Code = "Zips" Or Code = String.Empty Then
                strZips = "SELECT DISTINCT ZIP FROM tblSYS_ZIPCODES WHERE ZIPID >= 1" + _
                                IIf(pAddress.State.Trim <> "", " and STATE LIKE '%" + pAddress.State.Trim + "%'", "") + _
                                IIf(pAddress.City.Trim <> "", " AND CITY LIKE '%" + pAddress.City.Trim + "%'", "") + _
                                IIf(bolShowCounty And pAddress.County.Trim <> "", " AND COUNTY LIKE '%" + pAddress.County.Trim + "%'", "") + _
                                IIf(bolShowFIPS And pAddress.FIPSCode.Trim <> "", " AND FIPS LIKE '%" + pAddress.FIPSCode.Trim + "%'", "") + _
                                " Union Select ' ' as Zip " + _
                                " ORDER BY ZIP"
                cboZipCode.DataSource = pAddress.GetDataSet(strZips)
                If cboZipCode.Rows.Count = 2 Or cboZipCode.Rows.Count = 1 Then
                    pAddress.Zip = cboZipCode.Rows.Item(cboZipCode.Rows.Count - 1).Cells(0).Value
                End If
            End If


            If bolShowCounty AndAlso (Code = "Counties" Or Code = String.Empty) Then
                strCounties = "SELECT DISTINCT COUNTY FROM tblSYS_ZIPCODES WHERE ZIPID >= 1" + _
                                IIf(pAddress.State.Trim <> "", " and STATE LIKE '%" + pAddress.State.Trim + "%'", "") + _
                                IIf(pAddress.City.Trim <> "", " AND CITY LIKE '%" + pAddress.City.Trim + "%'", "") + _
                                IIf(pAddress.Zip.Trim <> "", " AND ZIP LIKE '%" + pAddress.Zip.Trim + "%'", "") + _
                                IIf(bolShowFIPS And pAddress.FIPSCode.Trim <> "", " AND FIPS LIKE '%" + pAddress.FIPSCode.Trim + "%'", "") + _
                               " Union Select ' ' as County " + _
                                " ORDER BY COUNTY"
                cboCounty.DataSource = pAddress.GetDataSet(strCounties)
                If cboCounty.Rows.Count = 2 Or cboCounty.Rows.Count = 1 Then
                    pAddress.County = cboCounty.Rows.Item(cboCounty.Rows.Count - 1).Cells(0).Value
                End If
            End If

            If bolShowFIPS AndAlso (Code = "FIPS" Or Code = String.Empty) Then
                strFips = "SELECT DISTINCT FIPS FROM tblSYS_ZIPCODES WHERE ZIPID >= 1 " + _
                                IIf(pAddress.State.Trim <> "", " and STATE LIKE '%" + pAddress.State.Trim + "%'", "") + _
                                IIf(pAddress.City.Trim <> "", " AND CITY LIKE '%" + pAddress.City.Trim + "%'", "") + _
                                IIf(bolShowCounty And pAddress.County.Trim <> "", " AND COUNTY LIKE '%" + pAddress.County.Trim + "%'", "") + _
                                IIf(pAddress.Zip.Trim <> "", " AND ZIP LIKE '%" + pAddress.Zip.Trim + "%'", "") + _
                                 " Union Select ' ' as FIPS " + _
                                " ORDER BY FIPS"
                cboFIPS.DataSource = pAddress.GetDataSet(strFips)
                If cboFIPS.Rows.Count = 1 Or cboFIPS.Rows.Count = 2 Then
                    pAddress.FIPSCode = cboFIPS.Rows.Item(cboFIPS.Rows.Count - 1).Cells(0).Value
                End If
            End If

            PopulateForm()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateForm()
        Try
            bolLoading = True
            txtAddress1.Text = pAddress.AddressLine1
            txtAddress2.Text = pAddress.AddressLine2
            TxtPhysicalTown.Text = pAddress.PhysicalTown
            cboState.Text = pAddress.State
            If ContainsValue(cboCity, pAddress.City) Then
                cboCity.Text = pAddress.City
            Else
                cboCity.Text = String.Empty
            End If
            If ContainsValue(cboZipCode, pAddress.Zip) Then
                cboZipCode.Text = pAddress.Zip
            Else
                cboZipCode.Text = String.Empty
            End If
            If ShowCounty Then
                If ContainsValue(cboCounty, pAddress.County) Then
                    cboCounty.Text = pAddress.County
                Else
                    cboCounty.Text = String.Empty
                End If
            End If
            If ShowFIPS Then
                If ContainsValue(cboFIPS, pAddress.FIPSCode) Then
                    cboFIPS.Text = pAddress.FIPSCode
                Else
                    cboFIPS.Text = String.Empty
                End If
            End If
            bolLoading = False


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function ContainsValue(ByVal cbo As Infragistics.Win.UltraWinGrid.UltraCombo, ByVal value As String) As Boolean
        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In cbo.Rows
            If row.Cells(0).Value = value Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Function ValidateData() As Boolean
        Try
            Dim strErr As String
            If pAddress.AddressLine1.Trim = String.Empty Then
                strErr += vbCrLf + lblAddress1.Text
            End If
            If bolShowFIPS Then
                If pAddress.FIPSCode.Trim = String.Empty Then
                    strErr += vbCrLf + lblFIPS.Text
                End If
            End If
            If bolShowCounty Then
                If pAddress.County.Trim = String.Empty Then
                    strErr += vbCrLf + lblCounty.Text
                End If
            End If
            If pAddress.City.Trim = String.Empty Then
                strErr += vbCrLf + lblCity.Text
            End If
            If pAddress.State.Trim = String.Empty Then
                strErr += vbCrLf + lblState.Text
            End If
            If pAddress.Zip.Trim = String.Empty Then
                strErr += vbCrLf + lblZip.Text
            End If
            If strErr <> String.Empty Then
                MsgBox("The following are required: " + strErr.TrimEnd(vbCrLf))
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "Exposed Properties"
    Public Property ShowCounty() As Boolean
        Get
            Return bolShowCounty
        End Get
        Set(ByVal Value As Boolean)
            bolShowCounty = Value
            lblCounty.Visible = Value
            cboCounty.Visible = Value
        End Set
    End Property
    Public Property ShowFIPS() As Boolean
        Get
            Return bolShowFIPS
        End Get
        Set(ByVal Value As Boolean)
            bolShowFIPS = Value
            lblFIPS.Visible = Value
            cboFIPS.Visible = Value
        End Set
    End Property
    Public Property ShowCopyFromOwner() As Boolean
        Get
            Return chkBoxCopyFromOwner.Visible
        End Get
        Set(ByVal Value As Boolean)
            chkBoxCopyFromOwner.Visible = Value
        End Set
    End Property
#End Region
#Region "UI Control Events"
    Private Sub btnAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAccept.Click
        Try
            If ValidateData() Then
                Dim oldID As Integer = pAddress.AddressId
                pAddress.EntityType = nEntityType
                If pAddress.AddressId <= 0 Then
                    pAddress.CreatedBy = MusterContainer.AppUser.ID
                Else
                    pAddress.ModifiedBy = MusterContainer.AppUser.ID
                End If
                Dim success As Boolean = False
                Try
                    success = pAddress.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, nEntityType, nEntityID, True)
                Catch ex As Exception

                End Try

                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                If success Then
                    If oldID <> pAddress.AddressId Then
                        RaiseEvent NewAddressID(pAddress.AddressId)
                    End If
                    MsgBox("Address saved")
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                Else
                    Me.DialogResult = DialogResult.Cancel
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        pAddress.Reset()
        Me.DialogResult = DialogResult.Cancel

        Me.Close()
    End Sub


    Private Sub cboCity_Leave2(ByVal sender As Object, ByVal e As EventArgs) Handles cboCity.TextChanged
        If Not bolLoading Then
            pAddress.City = cboCity.Text
        End If

        If pAddress.City.ToUpper = pAddress.PhysicalTown.ToUpper Then
            TxtPhysicalTown.Text = cboCity.Text
        End If

    End Sub

    Private Sub cboCity_Leave(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboCity.BeforeDropDown
        UpdateComboBoxes("Cities")
    End Sub

    Private Sub cboCounty_Leave2(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCounty.ValueChanged
        If Not bolLoading Then
            pAddress.County = cboCounty.Text
        End If
    End Sub
    Private Sub cboCounty_Leave(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboCounty.BeforeDropDown
        UpdateComboBoxes("Counties")
    End Sub


    Private Sub cboState_Leave2(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboState.TextChanged
        If Not bolLoading Then
            pAddress.State = cboState.Text
        End If

    End Sub

    Private Sub cboState_Leave(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboState.BeforeDropDown
        UpdateComboBoxes("States")
    End Sub

    Private Sub cboZipCode_Leave2(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboZipCode.TextChanged
        If Not bolLoading Then
            pAddress.Zip = cboZipCode.Text
        End If

    End Sub

    Private Sub cboZipCode_Leave(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboZipCode.BeforeDropDown
        UpdateComboBoxes("Zips")
    End Sub

    Private Sub TextPhysicaltown_Change(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtPhysicalTown.TextChanged
        If Me.TxtPhysicalTown.Text.ToUpper <> pAddress.PhysicalTown.ToUpper AndAlso Not bolLoading Then
            pAddress.PhysicalTown = Me.TxtPhysicalTown.Text
        End If
    End Sub

    Private Sub cboFIPS_Leave2(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFIPS.ValueChanged
        If Not bolLoading Then
            pAddress.FIPSCode = cboFIPS.Text
        End If

    End Sub

    Private Sub cboFIPS_Leave(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboFIPS.BeforeDropDown
        UpdateComboBoxes("FIPS")
    End Sub

    Private Sub txtAddress1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddress1.TextChanged
        If Not bolLoading Then
            pAddress.AddressLine1 = txtAddress1.Text.Trim
        End If
    End Sub

    Private Sub txtAddress2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddress2.TextChanged
        If Not bolLoading Then

            pAddress.AddressLine2 = txtAddress2.Text.Trim
        End If

    End Sub

    Private Sub btnClearData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearData.Click
        Try
            pAddress.FIPSCode = String.Empty
            pAddress.County = String.Empty
            pAddress.City = String.Empty
            pAddress.State = String.Empty
            pAddress.Zip = String.Empty
            UpdateComboBoxes(String.Empty)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub chkBoxCopyFromOwner_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBoxCopyFromOwner.CheckedChanged
        Try
            If chkBoxCopyFromOwner.Checked Then
                Dim addrLocal As New MUSTER.BusinessLogic.pAddress
                addrLocal.Retrieve(nownerAddressID)
                pAddress.AddressLine1 = addrLocal.AddressLine1.Trim
                pAddress.AddressLine2 = addrLocal.AddressLine2.Trim
                pAddress.City = addrLocal.City.Trim
                pAddress.County = addrLocal.County.Trim
                pAddress.FIPSCode = addrLocal.FIPSCode.Trim
                pAddress.State = addrLocal.State.Trim
                pAddress.Zip = addrLocal.Zip.Trim
                pAddress.PhysicalTown = addrLocal.PhysicalTown
                pAddress.Remove(nownerAddressID)
                addrLocal = Nothing
                UpdateComboBoxes(String.Empty)
            Else
                pAddress.AddressLine1 = String.Empty
                pAddress.AddressLine2 = String.Empty
                pAddress.PhysicalTown = String.Empty
                btnClearData.PerformClick()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Form Events"
    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        If pAddress.AddressId <= 0 Then
            pAddress.State = "MS"
        End If
        UpdateComboBoxes(String.Empty)
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        If pAddress.IsDirty Then
            pAddress.Reset()
        End If
    End Sub
#End Region
#Region "External Events"
    Private Sub AddressChanged(ByVal bolValue As Boolean) Handles pAddress.evtAddressChanged
        btnAccept.Enabled = bolValue
    End Sub
    Private Sub AddressesChanged(ByVal bolValue As Boolean) Handles pAddress.evtAddressesChanged
        btnAccept.Enabled = bolValue
    End Sub
    Private Sub AddressError(ByVal MsgStr As String) Handles pAddress.evtAddressErr
        MsgBox(MsgStr)
    End Sub
#End Region
#Region "ExceptionClasses"
    Public Class NoAddressException
        Inherits Exception

        Public Sub New()
            MyBase.New("Please enter in an address")
        End Sub

    End Class
#End Region

    Private Sub btnMap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMap.Click

        Try
            Dim longitude As Decimal = -1
            Dim latitude As Decimal = -1

            If Not _container.pOwn Is Nothing AndAlso Not _container.pOwn.Facilities Is Nothing Then
                With _container.pOwn.Facilities
                    If .LongitudeDegree > 0 AndAlso .LongitudeMinutes >= 0 AndAlso .LatitudeDegree > 0 AndAlso .LatitudeMinutes >= 0 AndAlso .LatitudeSeconds >= 0 AndAlso .LongitudeSeconds >= 0 Then
                        longitude = .LongitudeDegree + (.LongitudeMinutes / 60) + (IIf(.LongitudeSeconds >= 0, .LongitudeSeconds / 3600, 0))
                        latitude = .LatitudeDegree + (.LatitudeMinutes / 60) + (IIf(.LatitudeSeconds >= 0, .LatitudeSeconds / 3600, 0))
                    End If
                End With
            End If




            Dim mapAddress As New MapAddress(Me.txtAddress1.Text, Me.txtAddress2.Text, Me.TxtPhysicalTown.Text, Me.cboState.Text, longitude, latitude)
            mapAddress.ShowOnScreen()
            mapAddress.dispose()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
End Class
