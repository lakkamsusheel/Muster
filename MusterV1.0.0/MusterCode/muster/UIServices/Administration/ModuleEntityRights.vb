Public Class ModuleEntityRights
    Inherits System.Windows.Forms.Form

#Region "Private Member Variables"
    Private oUser As MUSTER.BusinessLogic.pUser

    Dim ugRow, ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private vListModule, vListEntity As Infragistics.Win.ValueList
    Dim rp As New Remove_Pencil

    Friend MyGUID As New System.Guid
    Dim returnVal As String = String.Empty
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef frm As Windows.Forms.Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MyGUID = System.Guid.NewGuid
        MusterContainer.AppUser.LogEntry(Me.Text, MyGUID.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGUID.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)
        If Not frm Is Nothing Then
            If frm.IsMdiContainer Then
                Me.MdiParent = frm
            End If
        End If
        oUser = New MUSTER.BusinessLogic.pUser
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
    Friend WithEvents ug As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ug = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        CType(Me.ug, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTop.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'ug
        '
        Me.ug.Cursor = System.Windows.Forms.Cursors.Default
        Me.ug.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ug.Location = New System.Drawing.Point(0, 0)
        Me.ug.Name = "ug"
        Me.ug.Size = New System.Drawing.Size(448, 317)
        Me.ug.TabIndex = 0
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.ug)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(448, 317)
        Me.pnlTop.TabIndex = 1
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.btnSave)
        Me.pnlBottom.Controls.Add(Me.btnDelete)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 317)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(448, 40)
        Me.pnlBottom.TabIndex = 1
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(143, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(231, 8)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.TabIndex = 0
        Me.btnDelete.Text = "Delete"
        '
        'ModuleEntityRights
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(448, 357)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.pnlBottom)
        Me.Name = "ModuleEntityRights"
        Me.Text = "ModuleEntityRights"
        CType(Me.ug, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTop.ResumeLayout(False)
        Me.pnlBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        LoadGrid()
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        '
        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGUID.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
        MusterContainer.AppUser.LogExit(MyGUID.ToString)
    End Sub

    Private Sub LoadGrid()
        ug.DataSource = oUser.ListModuleEntityRel
        ug.DrawFilter = rp
    End Sub
    Private Sub ShowError(ByVal ex As Exception)
        Dim MyErr As New ErrorReport(ex)
        MyErr.ShowDialog()
    End Sub

    Private Sub SetugRowComboValue(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            If ug.Band.Index = 0 Then
                If vListModule.FindByDataValue(ug.Cells("MODULE").Value) Is Nothing Then
                    ug.Cells("MODULE").Value = DBNull.Value
                End If
            ElseIf ug.Band.Index = 1 Then
                If vListEntity.FindByDataValue(ug.Cells("ENTITY").Value) Is Nothing Then
                    ug.Cells("ENTITY").Value = DBNull.Value
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ug_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ug.InitializeLayout
        Try
            e.Layout.Bands(0).Columns("MODULE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            e.Layout.Bands(1).Columns("ENTITY").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

            e.Layout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom

            e.Layout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False

            e.Layout.Bands(0).Columns("MODULE").Width = 200
            e.Layout.Bands(1).Columns("ENTITY").Width = 200

            e.Layout.Bands(1).Columns("MODULE").Hidden = True

            If e.Layout.Bands(0).Columns("MODULE").ValueList Is Nothing Then
                vListModule = New Infragistics.Win.ValueList
                For Each row As DataRow In oUser.ListPrimaryModules.Rows
                    vListModule.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(0).Columns("MODULE").ValueList = vListModule
            End If
            If e.Layout.Bands(1).Columns("ENTITY").ValueList Is Nothing Then
                vListEntity = New Infragistics.Win.ValueList
                For Each row As DataRow In oUser.ListEntityTypes.Rows
                    vListEntity.ValueListItems.Add(row.Item("ENTITY_ID"), row.Item("ENTITY_NAME").ToString)
                Next
                e.Layout.Bands(1).Columns("ENTITY").ValueList = vListEntity
            End If

            For Each ugRow In e.Layout.Grid.Rows
                SetugRowComboValue(ugRow)
            Next
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub ug_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ug.CellChange
        Try
            Dim newValue As Integer = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
            If e.Cell.Row.Band.Index = 0 Then
                For Each ugRow In ug.Rows
                    If Not ugRow.IsAddRow Then
                        If ugRow.Cells("MODULE").Value = newValue Then
                            MsgBox("Module already exists")
                            e.Cell.CancelUpdate()
                            Exit Sub
                        End If
                    End If
                Next
                e.Cell.Value = newValue
                e.Cell.Row.Tag = "1"
                If Not e.Cell.Row.ChildBands Is Nothing Then
                    If Not e.Cell.Row.ChildBands(0).Rows Is Nothing Then
                        If e.Cell.Row.ChildBands(0).Rows.Count > 0 Then
                            For Each ugChildRow In e.Cell.Row.ChildBands(0).Rows
                                ugChildRow.Cells("MODULE").Value = e.Cell.Value
                                ugChildRow.Tag = "1"
                            Next
                        End If
                    End If
                End If
            ElseIf e.Cell.Row.Band.Index = 1 Then
                If Not e.Cell.Row.ParentRow.ChildBands(0).Rows Is Nothing Then
                    For Each ugRow In e.Cell.Row.ParentRow.ChildBands(0).Rows
                        If Not ugRow.IsAddRow Then
                            If ugRow.Cells("ENTITY").Value = newValue Then
                                MsgBox("Entity already exists")
                                e.Cell.CancelUpdate()
                                Exit Sub
                            End If
                        End If
                    Next
                End If
                e.Cell.Value = newValue
                e.Cell.Row.Tag = "1"
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim bolSaved As Boolean = False
        Try
            For Each ugRow In ug.Rows
                If Not ugRow.ChildBands Is Nothing Then
                    If Not ugRow.ChildBands(0).Rows Is Nothing Then
                        For Each ugChildRow In ugRow.ChildBands(0).Rows
                            If Not ugChildRow.Tag Is Nothing Then
                                If ugChildRow.Tag = "1" Then
                                    oUser.SaveModuleEntityRel(ugChildRow.Cells("MODULE").Value, ugChildRow.Cells("ENTITY").Value, False)
                                    ugChildRow.Tag = "0"
                                    bolSaved = True
                                End If
                            End If
                        Next
                    End If
                End If
            Next
            If bolSaved Then
                MsgBox("Saved successfully")
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If Not ug.Selected.Rows Is Nothing Then
                If ug.Selected.Rows.Count > 0 Then
                    If MsgBox("Are you sure you want to delete the selected values?", MsgBoxStyle.YesNoCancel, "Delete Module Entity Relation") = MsgBoxResult.Yes Then
                        For Each ugRow In ug.Selected.Rows
                            If ugRow.Band.Index = 0 Then
                                If Not ugRow.ChildBands Is Nothing Then
                                    If Not ugRow.ChildBands(0).Rows Is Nothing Then
                                        For Each ugChildRow In ugRow.ChildBands(0).Rows
                                            If Not ugChildRow.IsAddRow Then
                                                oUser.SaveModuleEntityRel(ugChildRow.Cells("MODULE").Value, ugChildRow.Cells("ENTITY").Value, True)
                                            End If
                                        Next
                                    End If
                                End If
                            ElseIf ugRow.Band.Index = 1 Then
                                If Not ugRow.IsAddRow Then
                                    oUser.SaveModuleEntityRel(ugRow.Cells("MODULE").Value, ugRow.Cells("ENTITY").Value, True)
                                End If
                            End If
                        Next
                        LoadGrid()
                    End If
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

End Class
