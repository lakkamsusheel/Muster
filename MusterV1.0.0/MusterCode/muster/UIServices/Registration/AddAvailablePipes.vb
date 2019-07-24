' Release   Initials    Date        Description
'  2.0      TF          2/18/2009         Added Functionaility to LoadPipeList to access pipe and tank ststus columns



Public Class AddAvailablePipes
    Inherits System.Windows.Forms.Form
    Friend Facility_id As Integer = 0
    Friend Tank_id As Integer = 0
    Friend Compartment_number As Integer = 0
    Friend CallingForm As Form
    'P1 02/20/05 start
    Private LocalPpipe As MUSTER.BusinessLogic.pPipe
    Private oTank As MUSTER.BusinessLogic.pTank
    Dim returnVal As String = String.Empty
#Region " Windows Form Designer generated code "

    Public Sub New(ByRef pPipe As MUSTER.BusinessLogic.pPipe, ByRef pTank As MUSTER.BusinessLogic.pTank)
        MyBase.New()
        LocalPpipe = pPipe
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        oTank = pTank
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
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents lstViewAvailablPipes As System.Windows.Forms.ListView
    Friend WithEvents lsColPipeIndex As System.Windows.Forms.ColumnHeader
    Friend WithEvents lsColTankIndex As System.Windows.Forms.ColumnHeader
    Friend WithEvents lsColCompartmentNumber As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnAddExistingPipe As System.Windows.Forms.Button
    Friend WithEvents lsColSelect As System.Windows.Forms.ColumnHeader
    Friend WithEvents lblInstruction As System.Windows.Forms.Label
    Friend WithEvents btnDone As System.Windows.Forms.Button
    Friend WithEvents lsColPipeStatus As System.Windows.Forms.ColumnHeader
    Friend WithEvents lsColTankStatus As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.btnDone = New System.Windows.Forms.Button
        Me.lblInstruction = New System.Windows.Forms.Label
        Me.btnAddExistingPipe = New System.Windows.Forms.Button
        Me.lstViewAvailablPipes = New System.Windows.Forms.ListView
        Me.lsColSelect = New System.Windows.Forms.ColumnHeader
        Me.lsColPipeIndex = New System.Windows.Forms.ColumnHeader
        Me.lsColTankIndex = New System.Windows.Forms.ColumnHeader
        Me.lsColCompartmentNumber = New System.Windows.Forms.ColumnHeader
        Me.lsColPipeStatus = New System.Windows.Forms.ColumnHeader
        Me.lsColTankStatus = New System.Windows.Forms.ColumnHeader
        Me.pnlTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.btnDone)
        Me.pnlTop.Controls.Add(Me.lblInstruction)
        Me.pnlTop.Controls.Add(Me.btnAddExistingPipe)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(688, 64)
        Me.pnlTop.TabIndex = 0
        '
        'btnDone
        '
        Me.btnDone.Location = New System.Drawing.Point(200, 37)
        Me.btnDone.Name = "btnDone"
        Me.btnDone.TabIndex = 2
        Me.btnDone.Text = "Done"
        '
        'lblInstruction
        '
        Me.lblInstruction.AutoSize = True
        Me.lblInstruction.Location = New System.Drawing.Point(8, 16)
        Me.lblInstruction.Name = "lblInstruction"
        Me.lblInstruction.Size = New System.Drawing.Size(371, 16)
        Me.lblInstruction.TabIndex = 1
        Me.lblInstruction.Text = "Please Select a pipe from the list below  and Click  the  ""Add Pipe"" Button"
        '
        'btnAddExistingPipe
        '
        Me.btnAddExistingPipe.Location = New System.Drawing.Point(120, 37)
        Me.btnAddExistingPipe.Name = "btnAddExistingPipe"
        Me.btnAddExistingPipe.TabIndex = 0
        Me.btnAddExistingPipe.Text = "Attach Pipe"
        '
        'lstViewAvailablPipes
        '
        Me.lstViewAvailablPipes.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.lsColSelect, Me.lsColPipeIndex, Me.lsColPipeStatus, Me.lsColTankIndex, Me.lsColTankStatus, Me.lsColCompartmentNumber})
        Me.lstViewAvailablPipes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstViewAvailablPipes.FullRowSelect = True
        Me.lstViewAvailablPipes.GridLines = True
        Me.lstViewAvailablPipes.Location = New System.Drawing.Point(0, 64)
        Me.lstViewAvailablPipes.MultiSelect = False
        Me.lstViewAvailablPipes.Name = "lstViewAvailablPipes"
        Me.lstViewAvailablPipes.Size = New System.Drawing.Size(688, 430)
        Me.lstViewAvailablPipes.TabIndex = 1
        Me.lstViewAvailablPipes.View = System.Windows.Forms.View.Details
        '
        'lsColSelect
        '
        Me.lsColSelect.Text = ""
        Me.lsColSelect.Width = 33
        '
        'lsColPipeIndex
        '
        Me.lsColPipeIndex.Text = "Pipe Site ID"
        Me.lsColPipeIndex.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.lsColPipeIndex.Width = 69
        '
        'lsColTankIndex
        '
        Me.lsColTankIndex.Text = "Tank Site ID"
        Me.lsColTankIndex.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.lsColTankIndex.Width = 74
        '
        'lsColCompartmentNumber
        '
        Me.lsColCompartmentNumber.Text = "Compartment Number"
        Me.lsColCompartmentNumber.Width = 147
        '
        'lsColPipeStatus
        '
        Me.lsColPipeStatus.Text = "Pipe Status"
        Me.lsColPipeStatus.Width = 168
        '
        'lsColTankStatus
        '
        Me.lsColTankStatus.Text = "Tank Status"
        Me.lsColTankStatus.Width = 193
        '
        'AddAvailablePipes
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(688, 494)
        Me.Controls.Add(Me.lstViewAvailablPipes)
        Me.Controls.Add(Me.pnlTop)
        Me.Name = "AddAvailablePipes"
        Me.Text = "Add Existing Pipes"
        Me.pnlTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Function LoadPipeList()
        'Dim RDM As New InfoRepository.RegistrationDataManager
        Dim dsAvailablePipes As System.Data.DataTable
        Dim dRow As System.Data.DataRow
        Dim lstItem As New ListViewItem

        Try
            'dsAvailablePipes = RDM.getAvailablePipes(Facility_id, Tank_id, Compartment_number)
            dsAvailablePipes = LocalPpipe.ExistingPipesTable(Facility_id, Tank_id, Compartment_number)
            lstViewAvailablPipes.Items.Clear()
            For Each dRow In dsAvailablePipes.Rows

                If Convert.ToInt32(dRow("Has parent")) = 0 Then
                    lstItem = New ListViewItem
                    'lstItem.Tag = CType(dRow("PipeID"), Integer)
                    lstItem.Tag = CType(dRow("TankID"), Integer).ToString + "|" + CType(dRow("Compartment Number"), Integer).ToString + "|" + CType(dRow("PipeID"), Integer).ToString
                    'lstItem.SubItems.Add("")

                    If Not IsDBNull(dRow("Pipe Site ID")) Then
                        lstItem.SubItems.Add(CType(dRow("Pipe Site ID"), String))
                    Else
                        lstItem.SubItems.Add("")
                    End If

                    If Not IsDBNull(dRow("Pipe Site Status")) Then
                        lstItem.SubItems.Add(CType(dRow("Pipe Site Status"), String))
                    Else
                        lstItem.SubItems.Add("")
                    End If

                    If Not IsDBNull(dRow("Tank Site ID")) Then
                        lstItem.SubItems.Add(CType(dRow("Tank Site ID"), String))
                    Else
                        lstItem.SubItems.Add("")
                    End If


                    If Not IsDBNull(dRow("Tank Site Status")) Then
                        lstItem.SubItems.Add(CType(dRow("Tank Site Status"), String))
                    Else
                        lstItem.SubItems.Add("")
                    End If


                    If IsDBNull(dRow("Compartment Number")) Or IsDBNull(dRow("Compartment")) Then
                        lstItem.SubItems.Add("-NA-")
                    Else
                        If dRow("Compartment") Then
                            lstItem.SubItems.Add(CType(dRow("Compartment Number"), String))
                        Else
                            lstItem.SubItems.Add("-NA-")
                        End If
                    End If
                    lstViewAvailablPipes.Items.Add(lstItem)

                End If

            Next
            'MsgBox(dsAvailablePipes.Tables.Count.ToString)

        Catch ex As Exception
            Throw ex
        Finally
            ' RDM = Nothing
        End Try
    End Function

    Private Sub AddAvailablePipes_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LoadPipeList()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
            'MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub lstViewAvailablPipes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstViewAvailablPipes.Click
        Dim SelectedItem As ListViewItem
        Dim lstViewItem As ListViewItem
        Try
            For Each lstViewItem In lstViewAvailablPipes.Items
                lstViewItem.BackColor = Color.White
                lstViewItem.ForeColor = Color.Black
            Next
            If lstViewAvailablPipes.Items.Count <= 0 Or lstViewAvailablPipes.SelectedItems.Count <= 0 Then
                Exit Sub
            End If

            SelectedItem = lstViewAvailablPipes.SelectedItems(0)

            SelectedItem.BackColor = Color.Brown
            SelectedItem.ForeColor = Color.White
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
            'MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnAddExistingPipe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddExistingPipe.Click
        Dim lstViewItem As ListViewItem
        Dim boolFound As Boolean = False
        'Dim RDM As New InfoRepository.RegistrationDataManager
        'Dim pp As New InfoRepository.Pipe
        Dim reg As Registration
        Try
            For Each lstViewItem In lstViewAvailablPipes.Items
                If lstViewItem.BackColor.Equals(Color.Brown) And lstViewItem.ForeColor.Equals(Color.White) Then
                    LocalPpipe.Retrieve(oTank.TankInfo, 0, oTank.Compartments.CompInfo)
                    Dim strTag() As String = CType(lstViewItem.Tag, String).Split("|")
                    Dim nPipeID As Integer = CType(strTag(2), Integer)
                    LocalPpipe.ChangePipeTankCompartmentNumberKey(Tank_id, Compartment_number, nPipeID, LocalPpipe.Pipe)
                    LocalPpipe.CopyPipeInfo(CType(lstViewItem.Tag, String))
                    'RDM.AddPipeCompartment(pp)
                    If LocalPpipe.PipeID <= 0 Then
                        LocalPpipe.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        LocalPpipe.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    LocalPpipe.SaveCompartmentsPipe(CType(UIUtilsGen.ModuleID.Registration, Integer), MusterContainer.AppUser.UserKey, returnVal)


                    'Copies any extensions
                    If nPipeID > 0 Then

                        Dim pipeExtensions As DataTable = LocalPpipe.GetPipeExtensions(nPipeID)

                        If Not pipeExtensions Is Nothing AndAlso pipeExtensions.Rows.Count > 0 Then

                            For Each children As DataRow In pipeExtensions.Rows

                                Dim nChildPipeID As Integer = children("PIPE_ID")

                                LocalPpipe.Retrieve(oTank.TankInfo, 0, oTank.Compartments.CompInfo)
                                LocalPpipe.ChangePipeTankCompartmentNumberKey(Tank_id, Compartment_number, nChildPipeID, LocalPpipe.Pipe)
                                LocalPpipe.CopyPipeInfo(oTank.TankId.ToString)
                                'RDM.AddPipeCompartment(pp)
                                If LocalPpipe.PipeID <= 0 Then
                                    LocalPpipe.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    LocalPpipe.ModifiedBy = MusterContainer.AppUser.ID
                                End If
                                LocalPpipe.SaveCompartmentsPipe(CType(UIUtilsGen.ModuleID.Registration, Integer), MusterContainer.AppUser.UserKey, returnVal)

                            Next
                        End If


                    End If


                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    If CallingForm.GetType.ToString = GetType(Registration).ToString Then
                        reg = CType(CallingForm, Registration)
                        reg.PopulateTankPipeGrid(reg.nFacilityID, False)
                        reg.PopulateTank(Tank_id)
                        'reg.LoadTankData(Tank_id)
                        If reg.chkTankCompartment.Checked Then
                            MsgBox("Pipe successfully Attached to the Compartment")
                        Else
                            MsgBox("Pipe successfully Attached to the Tank")
                        End If
                    Else
                        MsgBox("Pipe successfully Attached to the Tank/Compartment")
                        CallingForm.Tag = "1"
                    End If
                    boolFound = True
                    Exit For
                End If
            Next

            If Not boolFound Then
                MsgBox("Please Select a Pipe from the List First", MsgBoxStyle.Exclamation)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
            'MsgBox(ex.Message)
        Finally
            'RDM = Nothing
            'pp = Nothing
            reg = Nothing
            AddAvailablePipes_Load(sender, e)
        End Try
    End Sub

    Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'P1 02/20/05 end
End Class
