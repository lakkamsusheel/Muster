Public Class ManageContactTypes
    Inherits System.Windows.Forms.Form
#Region "user defined variables"
    Dim pConStruct As New MUSTER.BusinessLogic.pContactStruct
    Dim moduleID As Integer
    Dim returnVal As String = String.Empty

#End Region
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
    Friend WithEvents ugContactTypes As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblNotes As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ugContactTypes = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.lblNotes = New System.Windows.Forms.Label
        CType(Me.ugContactTypes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ugContactTypes
        '
        Me.ugContactTypes.AllowDrop = True
        Me.ugContactTypes.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugContactTypes.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugContactTypes.Location = New System.Drawing.Point(16, 24)
        Me.ugContactTypes.Name = "ugContactTypes"
        Me.ugContactTypes.Size = New System.Drawing.Size(624, 280)
        Me.ugContactTypes.TabIndex = 193
        Me.ugContactTypes.Text = "CONTACT TYPES"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(664, 64)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 202
        Me.btnDelete.Text = "Delete"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(664, 104)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 26)
        Me.btnClose.TabIndex = 213
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(664, 24)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 26)
        Me.btnSave.TabIndex = 211
        Me.btnSave.Text = "Save"
        '
        'lblNotes
        '
        Me.lblNotes.Location = New System.Drawing.Point(16, 312)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Size = New System.Drawing.Size(624, 64)
        Me.lblNotes.TabIndex = 214
        Me.lblNotes.Text = "Note:"
        '
        'ManageContactTypes
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(760, 374)
        Me.Controls.Add(Me.lblNotes)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.ugContactTypes)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Name = "ManageContactTypes"
        Me.Text = "ManageContactTypes"
        CType(Me.ugContactTypes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ugContactTypes_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugContactTypes.InitializeLayout
        ugContactTypes.DisplayLayout.Bands(0).Columns("Module").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown

        ugContactTypes.DisplayLayout.Bands(0).Columns("ContactTypeID").Hidden = True
        ugContactTypes.DisplayLayout.Bands(0).Columns("ModuleID").Hidden = True
        ugContactTypes.DisplayLayout.Bands(0).Columns("deleted").Hidden = True
        ugContactTypes.DisplayLayout.Bands(0).Columns("LetterContactType").Hidden = True

        Dim vListModules As New Infragistics.Win.ValueList
        If ugContactTypes.DisplayLayout.Bands(0).Columns("Module").ValueList Is Nothing Then
            For Each row As DataRow In pConStruct.getModules.Tables(0).Rows
                vListModules.ValueListItems.Add(row.Item("ModuleID"), row.Item("ModuleName").ToString)
            Next
            ugContactTypes.DisplayLayout.Bands(0).Columns("Module").ValueList = vListModules
        Else
            vListModules = ugContactTypes.DisplayLayout.Bands(0).Columns("Module").ValueList
        End If

        Dim vListModules1 As New Infragistics.Win.ValueList
        If ugContactTypes.DisplayLayout.Bands(0).Columns("LetterContactTypeCode").ValueList Is Nothing Then
            For Each row As DataRow In pConStruct.getLetterContactType.Tables(0).Rows
                vListModules1.ValueListItems.Add(row.Item("LetterContactTypeID"), row.Item("LetterContactTypeName").ToString)
            Next
            ugContactTypes.DisplayLayout.Bands(0).Columns("LetterContactTypeCode").ValueList = vListModules1
        Else
            vListModules1 = ugContactTypes.DisplayLayout.Bands(0).Columns("LetterContactTypeCode").ValueList
        End If

        ' Set the appearance for template add-rows. Template add-rows are the 
        ' add-row templates that are displayed with each rows collection.
        '
        e.Layout.Override.TemplateAddRowAppearance.BackColor = Color.LightBlue
        e.Layout.Override.TemplateAddRowAppearance.ForeColor = Color.Yellow

        ' Once  the user modifies the contents of a template add-row, it becomes
        ' an add-row and the AddRowAppearance gets applied to such rows.
        '
        e.Layout.Override.AddRowAppearance.BackColor = Color.Yellow
        e.Layout.Override.AddRowAppearance.ForeColor = Color.Blue
        ugContactTypes.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom

        ugContactTypes.DisplayLayout.Bands(0).Columns("ContactType").Width = 300
        For Each drow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugContactTypes.Rows
            drow.Cells("Module").Value = drow.Cells("ModuleID").Value
            If vListModules.FindByDataValue(drow.Cells("Module").Value) Is Nothing Then
                drow.Cells("Module").Value = DBNull.Value
            End If
        Next
        For Each drow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugContactTypes.Rows
            drow.Cells("Module").Value = drow.Cells("ModuleID").Value
            If vListModules.FindByDataValue(drow.Cells("Module").Value) Is Nothing Then
                drow.Cells("Module").Value = DBNull.Value
            End If
        Next
        For Each drow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugContactTypes.Rows
            drow.Cells("LetterContactTypeCode").Value = drow.Cells("LetterContactType").Value
            If vListModules.FindByDataValue(drow.Cells("LetterContactTypeCode").Value) Is Nothing Then
                drow.Cells("LetterContactType").Value = DBNull.Value
            End If
        Next
    End Sub

    Private Sub ManageContactTypes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ugContactTypes.DataSource = pConStruct.getContactTypes()
            lblNotes.Text = "Notes: Only the selected row is saved upon clicking the save/delete button" + vbCrLf + _
                            " X  - This is a type that the module wants as an option" + vbCrLf + _
                            " XH - If there is contact of this type, the contact will be who the letter is addressed to rather than the registration owner" + vbCrLf + _
                            " XL - This contact will be the contact listed above the header address and in the letter Salutation."
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim id As Integer
        Dim LetterContactType As Integer
        Try
            If ugContactTypes.ActiveRow Is Nothing Then
                MsgBox("Please select the contactType row to be edited")
                Exit Sub
            End If
            If ugContactTypes.ActiveRow.Cells("ContactType").Value Is String.Empty Or _
            ugContactTypes.ActiveRow.Cells("Module").Value Is Nothing Or _
            ugContactTypes.ActiveRow.Cells("LetterContactTypeCode").Value Is Nothing Then
                MsgBox("ContactType, Module and LetterContactTypeCode are Required")
            End If
            id = ugContactTypes.ActiveRow.Cells("ContactTypeID").Value
            moduleID = ugContactTypes.ActiveRow.Cells("Module").Value
            LetterContactType = ugContactTypes.ActiveRow.Cells("LetterContactTypeCode").Value
            If ManageContactType(moduleID, LetterContactType) Then
                pConStruct.putContactType(ugContactTypes.ActiveRow.Cells("ContactType").Value, moduleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, id, , LetterContactType)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                MsgBox("Successfully saved")
                ugContactTypes.DataSource = pConStruct.getContactTypes()
            Else
                MsgBox("There's already a '" + ugContactTypes.ActiveRow.Cells("LetterContactTypeCode").Text + "' contact type for this module")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugContactTypes_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugContactTypes.AfterRowInsert
        ugContactTypes.ActiveRow.Cells("ContactTypeID").Value = 0
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub ugContactTypes_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugContactTypes.CellChange
        If e.Cell.Column.Key.ToUpper = "Module" Then
            'moduleID = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
            e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
        End If
    End Sub

    Private Sub ugContactTypes_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugContactTypes.AfterRowActivate
        'ugContactTypes.ActiveRow.Cells("Module").ValueList.
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim result As DialogResult
        Try
            result = MsgBox("Do you want to delete the selected contact type?", MsgBoxStyle.YesNoCancel)
            If result = DialogResult.Yes Then
                pConStruct.putContactType(ugContactTypes.ActiveRow.Cells("ContactType").Value, ugContactTypes.ActiveRow.Cells("Module").Value, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, ugContactTypes.ActiveRow.Cells("ContactTypeID").Value, 1)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                MsgBox("Successfuly deleted")
                ugContactTypes.DataSource = pConStruct.getContactTypes()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Public Function ManageContactType(ByVal moduleid As Integer, ByVal letterContactType As Integer) As Boolean
        Dim ds As DataSet
        Try
            ds = pConStruct.getContactTypes()
            For Each dr As DataRow In ds.Tables(0).Rows
                If dr("ModuleID") = moduleid And dr("LetterContactType") = letterContactType And Not dr("LetterContactType") = 1184 Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
End Class
