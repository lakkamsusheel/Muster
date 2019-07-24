Imports Infragistics
Public Class EditCompartments
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
    Friend WithEvents pnlNonCompProperties As System.Windows.Forms.Panel
    Friend WithEvents lblTankManifoldValue As System.Windows.Forms.Label
    Friend WithEvents cmbTankFuelType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbTanksubstance As System.Windows.Forms.ComboBox
    Friend WithEvents lblNonCompTankCapacity As System.Windows.Forms.Label
    Friend WithEvents lblTankFuelType As System.Windows.Forms.Label
    Friend WithEvents txtNonCompTankCapacity As System.Windows.Forms.TextBox
    Friend WithEvents lblTankSubstance As System.Windows.Forms.Label
    Friend WithEvents lblTankManifold As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlNonCompProperties = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnAdd = New System.Windows.Forms.Button
        Me.lblTankManifoldValue = New System.Windows.Forms.Label
        Me.cmbTankFuelType = New System.Windows.Forms.ComboBox
        Me.cmbTanksubstance = New System.Windows.Forms.ComboBox
        Me.lblNonCompTankCapacity = New System.Windows.Forms.Label
        Me.lblTankFuelType = New System.Windows.Forms.Label
        Me.txtNonCompTankCapacity = New System.Windows.Forms.TextBox
        Me.lblTankSubstance = New System.Windows.Forms.Label
        Me.lblTankManifold = New System.Windows.Forms.Label
        Me.pnlNonCompProperties.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlNonCompProperties
        '
        Me.pnlNonCompProperties.BackColor = System.Drawing.SystemColors.ControlLight
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankManifoldValue)
        Me.pnlNonCompProperties.Controls.Add(Me.cmbTankFuelType)
        Me.pnlNonCompProperties.Controls.Add(Me.cmbTanksubstance)
        Me.pnlNonCompProperties.Controls.Add(Me.lblNonCompTankCapacity)
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankFuelType)
        Me.pnlNonCompProperties.Controls.Add(Me.txtNonCompTankCapacity)
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankSubstance)
        Me.pnlNonCompProperties.Controls.Add(Me.lblTankManifold)
        Me.pnlNonCompProperties.Location = New System.Drawing.Point(8, 8)
        Me.pnlNonCompProperties.Name = "pnlNonCompProperties"
        Me.pnlNonCompProperties.Size = New System.Drawing.Size(744, 80)
        Me.pnlNonCompProperties.TabIndex = 4
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.Location = New System.Drawing.Point(688, 96)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(64, 24)
        Me.btnCancel.TabIndex = 209
        Me.btnCancel.Text = "Cancel"
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.SystemColors.Control
        Me.btnAdd.Location = New System.Drawing.Point(560, 96)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.TabIndex = 208
        Me.btnAdd.Text = "Add"
        '
        'lblTankManifoldValue
        '
        Me.lblTankManifoldValue.AutoSize = True
        Me.lblTankManifoldValue.Location = New System.Drawing.Point(120, 80)
        Me.lblTankManifoldValue.Name = "lblTankManifoldValue"
        Me.lblTankManifoldValue.Size = New System.Drawing.Size(0, 16)
        Me.lblTankManifoldValue.TabIndex = 207
        '
        'cmbTankFuelType
        '
        Me.cmbTankFuelType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTankFuelType.Location = New System.Drawing.Point(424, 16)
        Me.cmbTankFuelType.Name = "cmbTankFuelType"
        Me.cmbTankFuelType.Size = New System.Drawing.Size(208, 21)
        Me.cmbTankFuelType.TabIndex = 2
        '
        'cmbTanksubstance
        '
        Me.cmbTanksubstance.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTanksubstance.DropDownWidth = 250
        Me.cmbTanksubstance.Location = New System.Drawing.Point(120, 45)
        Me.cmbTanksubstance.Name = "cmbTanksubstance"
        Me.cmbTanksubstance.Size = New System.Drawing.Size(208, 21)
        Me.cmbTanksubstance.TabIndex = 1
        '
        'lblNonCompTankCapacity
        '
        Me.lblNonCompTankCapacity.Location = New System.Drawing.Point(24, 16)
        Me.lblNonCompTankCapacity.Name = "lblNonCompTankCapacity"
        Me.lblNonCompTankCapacity.Size = New System.Drawing.Size(88, 23)
        Me.lblNonCompTankCapacity.TabIndex = 197
        Me.lblNonCompTankCapacity.Text = "Tank Capacity:"
        '
        'lblTankFuelType
        '
        Me.lblTankFuelType.Location = New System.Drawing.Point(352, 16)
        Me.lblTankFuelType.Name = "lblTankFuelType"
        Me.lblTankFuelType.Size = New System.Drawing.Size(64, 23)
        Me.lblTankFuelType.TabIndex = 201
        Me.lblTankFuelType.Text = "Fuel Type: "
        '
        'txtNonCompTankCapacity
        '
        Me.txtNonCompTankCapacity.Location = New System.Drawing.Point(120, 16)
        Me.txtNonCompTankCapacity.Name = "txtNonCompTankCapacity"
        Me.txtNonCompTankCapacity.Size = New System.Drawing.Size(80, 20)
        Me.txtNonCompTankCapacity.TabIndex = 0
        Me.txtNonCompTankCapacity.Text = ""
        '
        'lblTankSubstance
        '
        Me.lblTankSubstance.Location = New System.Drawing.Point(40, 46)
        Me.lblTankSubstance.Name = "lblTankSubstance"
        Me.lblTankSubstance.Size = New System.Drawing.Size(72, 23)
        Me.lblTankSubstance.TabIndex = 203
        Me.lblTankSubstance.Text = "Substance: "
        '
        'lblTankManifold
        '
        Me.lblTankManifold.Location = New System.Drawing.Point(56, 80)
        Me.lblTankManifold.Name = "lblTankManifold"
        Me.lblTankManifold.Size = New System.Drawing.Size(56, 23)
        Me.lblTankManifold.TabIndex = 199
        Me.lblTankManifold.Text = "Manifold:"
        '
        'EditCompartments
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(760, 126)
        Me.Controls.Add(Me.pnlNonCompProperties)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.btnCancel)
        Me.Name = "EditCompartments"
        Me.Text = "Adding New Compartment"
        Me.pnlNonCompProperties.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "private members"

    Private bIsNew As Boolean = True
    Private bRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private WithEvents dGridBand As Infragistics.Win.UltraWinGrid.UltraGridBand
    Private WithEvents dGrid As Infragistics.Win.UltraWinGrid.UltraGrid
    Private pTank As BusinessLogic.pTank
    Private bDelete As Boolean = False
#End Region

#Region "Properties"

    Private Property IsNew() As Boolean

        Get
            Return bIsNew
        End Get

        Set(ByVal Value As Boolean)
            bIsNew = Value

            If Not bIsNew Then
                Text = String.Format("Editing Compartment #", bRow.Cells("COMPARTMENT #"))
                btnAdd.Text = "Update"

            Else
                Text = "Adding New Compartment"
                btnAdd.Text = "Add"
            End If

        End Set

    End Property

#End Region

#Region "Construct"
    Sub New(ByVal tank As BusinessLogic.pTank, ByVal grid As Infragistics.Win.UltraWinGrid.UltraGrid, Optional ByVal tankRow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing, Optional ByVal deleteMe As Boolean = False)

        Me.New()

        pTank = tank
        bRow = tankRow
        dGridBand = grid.DisplayLayout.Bands(0)
        dGrid = grid
        bDelete = deleteMe

    End Sub


#End Region

#Region "Methods"

    Private Sub ShowError(ByVal ex As Exception)
        Dim MyErr As New ErrorReport(ex)
        MyErr.ShowDialog()
    End Sub

    Private Sub PopulateCompartmentSubstance()
        Try
            cmbTanksubstance.DataSource = pTank.PopulateCompartmentSubstance
            cmbTanksubstance.DisplayMember = "PROPERTY_NAME"
            cmbTanksubstance.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub PopulateCompartmentFuelType(Optional ByVal compSubstance As Int64 = 0)
        Try
            cmbTankFuelType.DataSource = pTank.PopulateCompartmentFuelTypes(compSubstance)
            cmbTankFuelType.DisplayMember = "PROPERTY_NAME"
            cmbTankFuelType.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub


    Private Sub FillForm()

        'Populate Lists

        If Not bRow Is Nothing Then

            PopulateCompartmentSubstance()

            cmbTanksubstance.SelectedValue = bRow.Cells("SUBSTANCE").Value

            PopulateCompartmentFuelType(cmbTanksubstance.SelectedValue)

            cmbTankFuelType.SelectedValue = bRow.Cells("FUEL TYPE ID").Value

            txtNonCompTankCapacity.Text = bRow.Cells("CAPACITY").Value.ToString
        Else

            PopulateCompartmentSubstance()
            PopulateCompartmentFuelType(cmbTanksubstance.SelectedValue)


        End If

    End Sub


    Private Sub SetugRowComboValue(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Dim vListFuelType As Infragistics.Win.ValueList
        Try
            ' Substance
            vListFuelType = New Infragistics.Win.ValueList
                Dim dt As DataTable = pTank.PopulateCompartmentFuelTypes(0)
                If dt Is Nothing Then
                    ug.Cells("FUEL TYPE ID").Value = DBNull.Value
                    ug.Cells("FUEL TYPE ID").Hidden = True
                Else
                    If dt.Rows.Count > 0 Then


                        ug.Cells("FUEL TYPE ID").Hidden = False

                        For Each row As DataRow In dt.Rows
                            vListFuelType.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                        Next
                        vListFuelType.ValueListItems.Add(-1, "N/A")

                        ug.Cells("FUEL TYPE ID").ValueList = vListFuelType
                        If vListFuelType.FindByDataValue(ug.Cells("FUEL TYPE ID").Value) Is Nothing Then
                            ug.Cells("FUEL TYPE ID").Value = DBNull.Value
                        End If
                    Else
                        ug.Cells("FUEL TYPE ID").Value = DBNull.Value
                        ug.Cells("FUEL TYPE ID").Hidden = True
                    End If
                End If

            ' Fuel Type
            'If vListFuelType.FindByDataValue(ug.Cells("FUEL TYPE ID").Value) Is Nothing Then
            '    ug.Cells("FUEL TYPE ID").Value = DBNull.Value
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SaveData()

        Dim capacity As Integer = 0

        Try
            capacity = IIf(txtNonCompTankCapacity.Text = "", 0, Convert.ToInt32(txtNonCompTankCapacity.Text))
        Catch ex As Exception
            MsgBox("Please Enter in a numeric value for capacity")
            Return
        End Try

        If IsNew Then

            With dGridBand

                .Override.AllowAddNew = Win.UltraWinGrid.AllowAddNew.Yes

                bRow = .AddNew

                SetugRowComboValue(bRow)
            End With

        End If

        With bRow

            .Band.Override.AllowUpdate = Win.DefaultableBoolean.True

            If Not Me.cmbTanksubstance.SelectedValue Is Nothing Then
                .Cells("SUBSTANCE").Value = cmbTanksubstance.SelectedValue
            Else
                .Cells("SUBSTANCE").Value = -1

            End If


            If Not Me.cmbTankFuelType.SelectedValue Is Nothing Then

                .Cells("FUEL TYPE ID").Value = Me.cmbTankFuelType.SelectedValue


            Else
                .Cells("FUEL TYPE ID").Value = -1

            End If

            .Cells("CAPACITY").Value = capacity

            .Refresh()


            .Band.Override.AllowUpdate = Win.DefaultableBoolean.False



        End With

        dGridBand.Override.AllowAddNew = Win.UltraWinGrid.AllowAddNew.No

        Me.Close()

    End Sub


    Private Sub DeleteData()



        Try
            With bRow



                Dim ID As Integer = .Cells("COMPARTMENT NUMBER").Value

                pTank.Compartments.Remove(pTank.TankId.ToString + "|" + ID.ToString)

                bRow.Delete()

                bRow.Refresh()

            End With
        Catch ex As Exception
            MsgBox("Please make sure a proper compartment row is selected before deleting. ")
        End Try

        Me.Close()

    End Sub


#End Region


#Region "Events"

    Private Sub dGridCompartments_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
        Try


            If "CAPACITY".Equals(e.Cell.Column.Key) Then
                If pTank.Compartments.COMPARTMENTNumber <> e.Cell.Row.Cells("COMPARTMENT NUMBER").Value Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, pTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT NUMBER").Value.ToString, False)
                End If
                pTank.Compartments.Capacity = e.Cell.Value

            ElseIf "SUBSTANCE".Equals(e.Cell.Column.Key) Then
                If pTank.Compartments.COMPARTMENTNumber <> e.Cell.Row.Cells("COMPARTMENT NUMBER").Value Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, pTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT NUMBER").Value.ToString, False)
                End If
                pTank.Compartments.Substance = IIf(e.Cell.Value Is Nothing, -1, e.Cell.Value)
            ElseIf "FUEL TYPE ID".Equals(e.Cell.Column.Key) Then
                If pTank.Compartments.COMPARTMENTNumber <> e.Cell.Row.Cells("COMPARTMENT NUMBER").Value Then
                    pTank.Compartments.Retrieve(pTank.TankInfo, pTank.TankId.ToString + "|" + e.Cell.Row.Cells("COMPARTMENT NUMBER").Value.ToString, False)
                End If
                pTank.Compartments.FuelTypeId = IIf(e.Cell.Value Is Nothing, -1, e.Cell.Value)


            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub EditCompartments_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Not bRow Is Nothing Then
            IsNew = False
        End If

        If Not bDelete Then
            FillForm()

            AddHandler dGrid.AfterCellUpdate, AddressOf dGridCompartments_CellChange

        Else

            Me.Controls.Remove(pnlNonCompProperties)

            btnAdd.Text = "Yes"
            btnCancel.Text = "No"
            Me.Text = "Delete Compartment"

            Dim lbl As New Label
            lbl.Font = New Font(Drawing.FontFamily.GenericSansSerif, 10, FontStyle.Bold, GraphicsUnit.Point, 10)
            lbl.Width = 1000
            lbl.AutoSize = False
            lbl.Height = 50


            lbl.Text = String.Format("Do you want to remove Compartment {0} from Tank  #{1}? ", bRow.Cells("COMPARTMENT #").Value, pTank.TankIndex)

            Controls.Add(lbl)

        End If

    End Sub

    Public Sub SaveClicked(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        If bDelete Then

            DeleteData()

        Else
            SaveData()

        End If
    End Sub

    Public Sub CloseClicked(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub cmbTanksubstance_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTanksubstance.SelectedValueChanged

        If TypeOf cmbTanksubstance.SelectedValue Is Integer Then
            PopulateCompartmentFuelType(cmbTanksubstance.SelectedValue)
        End If

    End Sub

#End Region

    Private Sub EditCompartments_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        RemoveHandler cmbTanksubstance.SelectedValueChanged, AddressOf cmbTanksubstance_SelectedValueChanged
        RemoveHandler dGrid.AfterCellUpdate, AddressOf dGridCompartments_CellChange

    End Sub
End Class


