Public Class ContactReconciliation
    Inherits System.Windows.Forms.Form
    Private pConStruct As MUSTER.BusinessLogic.pContactStruct
    Private dsRecon As DataSet
    Private ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private bolLoading As Boolean = False
    Private bolLoadFromLetters As Boolean = False
    Private nEntityID As Integer
    Private nEntityType As Integer
    Private nModuleID As Integer
    Private dtReconContactFromLetters As DataTable
    Dim returnVal As String = String.Empty

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef pContStruct As MUSTER.BusinessLogic.pContactStruct, Optional ByVal LoadFromLetters As Boolean = False, Optional ByVal dtReconciliation As DataTable = Nothing, Optional ByVal nEntID As Integer = 0, Optional ByVal nEntType As Integer = 0, Optional ByVal nModID As Integer = 0)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        pConStruct = pContStruct
        bolLoadFromLetters = LoadFromLetters
        If bolLoadFromLetters Then
            dtReconContactFromLetters = dtReconciliation
            nEntityID = nEntID
            nEntityType = nEntType
            nModuleID = nModID
        End If
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
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents ugContactReconciliation As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnRejectAll As System.Windows.Forms.Button
    Friend WithEvents btnAcceptAll As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.ugContactReconciliation = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnRejectAll = New System.Windows.Forms.Button
        Me.btnAcceptAll = New System.Windows.Forms.Button
        Me.btnClear = New System.Windows.Forms.Button
        CType(Me.ugContactReconciliation, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(512, 464)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(70, 23)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(432, 464)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(70, 23)
        Me.btnOk.TabIndex = 4
        Me.btnOk.Text = "Ok"
        '
        'ugContactReconciliation
        '
        Me.ugContactReconciliation.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugContactReconciliation.Location = New System.Drawing.Point(8, 48)
        Me.ugContactReconciliation.Name = "ugContactReconciliation"
        Me.ugContactReconciliation.Size = New System.Drawing.Size(984, 400)
        Me.ugContactReconciliation.TabIndex = 3
        '
        'btnRejectAll
        '
        Me.btnRejectAll.Location = New System.Drawing.Point(840, 16)
        Me.btnRejectAll.Name = "btnRejectAll"
        Me.btnRejectAll.Size = New System.Drawing.Size(70, 23)
        Me.btnRejectAll.TabIndex = 1
        Me.btnRejectAll.Text = "Reject All"
        '
        'btnAcceptAll
        '
        Me.btnAcceptAll.Location = New System.Drawing.Point(760, 16)
        Me.btnAcceptAll.Name = "btnAcceptAll"
        Me.btnAcceptAll.Size = New System.Drawing.Size(70, 23)
        Me.btnAcceptAll.TabIndex = 0
        Me.btnAcceptAll.Text = "Accept All"
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(920, 16)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(70, 23)
        Me.btnClear.TabIndex = 2
        Me.btnClear.Text = "Clear"
        '
        'ContactReconciliation
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1028, 541)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnRejectAll)
        Me.Controls.Add(Me.btnAcceptAll)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.ugContactReconciliation)
        Me.Name = "ContactReconciliation"
        Me.Text = "Contact Reconciliation"
        CType(Me.ugContactReconciliation, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ContactReconciliation_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True
            populateReconciliation()
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Try
            Dim strAccept As String = String.Empty
            Dim strReject As String = String.Empty
            Dim bolStatus As Boolean = False

            For Each ugrow In ugContactReconciliation.Rows

                If ugrow.Cells("ACCEPT").Value Then
                    bolStatus = True
                    If strAccept <> String.Empty Then
                        strAccept += "," + ugrow.Cells("RECONCILIATIONID").Value.ToString
                    Else
                        strAccept += ugrow.Cells("RECONCILIATIONID").Value.ToString
                    End If
                ElseIf ugrow.Cells("REJECT").Value Then
                    bolStatus = True
                    If strReject <> String.Empty Then
                        strReject += "," + ugrow.Cells("RECONCILIATIONID").Value.ToString
                    Else
                        strReject += ugrow.Cells("RECONCILIATIONID").Value.ToString
                    End If
                End If
            Next

            If bolStatus Then
                ' TO DO: there is a possibility of accepting and also rejecting the same contact - this shouldnt occur-in this case exception occurs
                pConStruct.UpdateReconciliation(strAccept, strReject, CType(UIUtilsGen.ModuleID.ContactManagement, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                MsgBox("Address Reconciliation was Successfull.")
                If bolLoadFromLetters Then
                    bolLoadFromLetters = False
                    Me.Close()
                End If
                'ugContactReconciliation.DataSource = pConStruct.getReconciliation.Tables(0).DefaultView
                populateReconciliation()
            Else
                MsgBox("No Records Found OR Check at least one Accept or Reject Option")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try

            Me.Close()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnAcceptAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAcceptAll.Click
        Try
            bolLoading = True
            For Each ugrow In ugContactReconciliation.Rows
                ugrow.Cells("ACCEPT").Value = True
                ugrow.Cells("REJECT").Value = False
            Next
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnRejectAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRejectAll.Click
        Try
            bolLoading = True
            For Each ugrow In ugContactReconciliation.Rows
                ugrow.Cells("REJECT").Value = True
                ugrow.Cells("ACCEPT").Value = False
            Next
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugContactReconciliation_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugContactReconciliation.AfterCellUpdate
        Try
            If bolLoading Then Exit Sub

            If e.Cell.Column.ToString.ToUpper = "ACCEPT".ToUpper And e.Cell.Value Then
                If e.Cell.Row.Cells("REJECT").Value Then
                    e.Cell.Value = e.Cell.OriginalValue
                End If
            End If

            If e.Cell.Column.ToString.ToUpper = "REJECT".ToUpper And e.Cell.Value Then
                If e.Cell.Row.Cells("ACCEPT").Value Then
                    bolLoading = True
                    e.Cell.Value = False
                    bolLoading = False
                End If
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Try
            For Each ugrow In ugContactReconciliation.Rows
                ugrow.Cells("ACCEPT").Value = False
                ugrow.Cells("REJECT").Value = False
            Next
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub ugContactReconciliation_AfterEnterEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugContactReconciliation.AfterEnterEditMode
        Dim MyCell As Infragistics.Win.UltraWinGrid.UltraGridCell

        Try
            MyCell = ugContactReconciliation.ActiveCell
            If Not (MyCell.Column.ToString.ToUpper = "ACCEPT".ToUpper Or MyCell.Column.ToString.ToUpper = "REJECT".ToUpper) Then
                ugContactReconciliation.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub ugContactReconciliation_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugContactReconciliation.InitializeLayout
    '    ugContactReconciliation.DisplayLayout.Bands(0).Columns("RECONCILIATIONID").Hidden = True
    'End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            If bolLoadFromLetters Then
                MsgBox("Either Accept or Reject option has to be selected")
                e.Cancel = True
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub populateReconciliation()
        If Not bolLoadFromLetters Then
            dsRecon = pConStruct.getReconciliation()
            ugContactReconciliation.DataSource = dsRecon.Tables(0).DefaultView
        Else
            dtReconContactFromLetters.DefaultView.RowFilter = "entityid= " + nEntityID.ToString + " and entitytype=" + nEntityType.ToString + " and moduleid=" + nModuleID.ToString
            ugContactReconciliation.DataSource = dtReconContactFromLetters.DefaultView
        End If
        ugContactReconciliation.DisplayLayout.Bands(0).Columns("EntityID").Hidden = True
        ugContactReconciliation.DisplayLayout.Bands(0).Columns("EntityType").Hidden = True
        ugContactReconciliation.DisplayLayout.Bands(0).Columns("ModuleID").Hidden = True
        ugContactReconciliation.DisplayLayout.Bands(0).Columns("RECONCILIATIONID").Hidden = True
        If Not bolLoadFromLetters Then
            ugContactReconciliation.DisplayLayout.Bands(0).Columns("ENTITYASSOCID").Hidden = True
            ugContactReconciliation.DisplayLayout.Bands(0).Columns("OLD_ContactAssocID").Hidden = True
            ugContactReconciliation.DisplayLayout.Bands(0).Columns("NEW_CONTACTASSOCID").Hidden = True
            ugContactReconciliation.DisplayLayout.Bands(0).Columns("Modified_EntityAssocID").Hidden = True
        End If
        ugContactReconciliation.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
    End Sub

    Private Sub ugContactReconciliation_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugContactReconciliation.InitializeLayout

    End Sub
End Class
