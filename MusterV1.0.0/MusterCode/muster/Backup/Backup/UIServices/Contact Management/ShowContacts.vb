Public Class ShowContacts
    Inherits System.Windows.Forms.Form
    Private WithEvents ContactFrm As Contacts
    Dim pConStruct As MUSTER.BusinessLogic.pContactStruct
    Dim nEntityID As Integer = 0
    Dim dsContacts As DataSet
    Dim nModuleID As Integer = 0
    Dim result As DialogResult
    Dim returnVal As String = String.Empty
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        nModuleID = 612
        nEntityID = "2007358"
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
    Friend WithEvents ugContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnAddSearchContact As System.Windows.Forms.Button
    Friend WithEvents lblContacts As System.Windows.Forms.Label
    Friend WithEvents chkShowActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ugContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnModifyContact = New System.Windows.Forms.Button
        Me.btnDeleteContact = New System.Windows.Forms.Button
        Me.btnAssociateContact = New System.Windows.Forms.Button
        Me.btnAddSearchContact = New System.Windows.Forms.Button
        Me.lblContacts = New System.Windows.Forms.Label
        Me.chkShowActiveOnly = New System.Windows.Forms.CheckBox
        Me.chkShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkShowContactsforAllModules = New System.Windows.Forms.CheckBox
        Me.btnClose = New System.Windows.Forms.Button
        CType(Me.ugContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ugContacts
        '
        Me.ugContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugContacts.Location = New System.Drawing.Point(8, 80)
        Me.ugContacts.Name = "ugContacts"
        Me.ugContacts.Size = New System.Drawing.Size(744, 240)
        Me.ugContacts.TabIndex = 3
        '
        'btnModifyContact
        '
        Me.btnModifyContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnModifyContact.Location = New System.Drawing.Point(304, 336)
        Me.btnModifyContact.Name = "btnModifyContact"
        Me.btnModifyContact.Size = New System.Drawing.Size(120, 23)
        Me.btnModifyContact.TabIndex = 5
        Me.btnModifyContact.Text = "Modify Contact"
        '
        'btnDeleteContact
        '
        Me.btnDeleteContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteContact.Location = New System.Drawing.Point(432, 336)
        Me.btnDeleteContact.Name = "btnDeleteContact"
        Me.btnDeleteContact.Size = New System.Drawing.Size(112, 23)
        Me.btnDeleteContact.TabIndex = 6
        Me.btnDeleteContact.Text = "Delete Contact"
        '
        'btnAssociateContact
        '
        Me.btnAssociateContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAssociateContact.Location = New System.Drawing.Point(552, 336)
        Me.btnAssociateContact.Name = "btnAssociateContact"
        Me.btnAssociateContact.Size = New System.Drawing.Size(128, 23)
        Me.btnAssociateContact.TabIndex = 7
        Me.btnAssociateContact.Text = "Associate Contact"
        '
        'btnAddSearchContact
        '
        Me.btnAddSearchContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddSearchContact.Location = New System.Drawing.Point(160, 336)
        Me.btnAddSearchContact.Name = "btnAddSearchContact"
        Me.btnAddSearchContact.Size = New System.Drawing.Size(136, 23)
        Me.btnAddSearchContact.TabIndex = 4
        Me.btnAddSearchContact.Text = "Add/Search Contact"
        '
        'lblContacts
        '
        Me.lblContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContacts.Location = New System.Drawing.Point(16, 56)
        Me.lblContacts.Name = "lblContacts"
        Me.lblContacts.Size = New System.Drawing.Size(56, 16)
        Me.lblContacts.TabIndex = 130
        Me.lblContacts.Text = "Contacts"
        '
        'chkShowActiveOnly
        '
        Me.chkShowActiveOnly.Checked = True
        Me.chkShowActiveOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowActiveOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowActiveOnly.Location = New System.Drawing.Point(552, 56)
        Me.chkShowActiveOnly.Name = "chkShowActiveOnly"
        Me.chkShowActiveOnly.Size = New System.Drawing.Size(144, 16)
        Me.chkShowActiveOnly.TabIndex = 2
        Me.chkShowActiveOnly.Tag = "646"
        Me.chkShowActiveOnly.Text = "Show Active Only"
        '
        'chkShowRelatedContacts
        '
        Me.chkShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowRelatedContacts.Location = New System.Drawing.Point(552, 32)
        Me.chkShowRelatedContacts.Name = "chkShowRelatedContacts"
        Me.chkShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkShowRelatedContacts.TabIndex = 1
        Me.chkShowRelatedContacts.Tag = "645"
        Me.chkShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkShowContactsforAllModules
        '
        Me.chkShowContactsforAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowContactsforAllModules.Location = New System.Drawing.Point(552, 8)
        Me.chkShowContactsforAllModules.Name = "chkShowContactsforAllModules"
        Me.chkShowContactsforAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkShowContactsforAllModules.TabIndex = 0
        Me.chkShowContactsforAllModules.Tag = "644"
        Me.chkShowContactsforAllModules.Text = "Show Contacts for All Modules"
        '
        'btnClose
        '
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(688, 336)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(64, 23)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "Close"
        '
        'ShowContacts
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(776, 374)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.chkShowActiveOnly)
        Me.Controls.Add(Me.chkShowRelatedContacts)
        Me.Controls.Add(Me.chkShowContactsforAllModules)
        Me.Controls.Add(Me.lblContacts)
        Me.Controls.Add(Me.ugContacts)
        Me.Controls.Add(Me.btnModifyContact)
        Me.Controls.Add(Me.btnDeleteContact)
        Me.Controls.Add(Me.btnAssociateContact)
        Me.Controls.Add(Me.btnAddSearchContact)
        Me.Name = "ShowContacts"
        Me.Text = "Show Contacts"
        CType(Me.ugContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub btnAddSearchContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddSearchContact.Click
        Try
            Dim objCntSearch As ContactSearch
            objCntSearch = New ContactSearch(nEntityID, 9, "Registration", pConStruct)
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
      
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ShowContacts_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct

            dsContacts = pConStruct.GetAll()
            dsContacts.Tables(0).DefaultView.RowFilter = "MODULEID = " + nModuleID.ToString + " And ENTITYID = " + nEntityID.ToString
            'dsContacts.Tables(0).DefaultView.Sort = "CONTACT_NAME ASC"
            ugContacts.DataSource = dsContacts.Tables(0).DefaultView
            ugContacts.DisplayLayout.Bands(0).Columns("Parent_Contact").Hidden = True
            ugContacts.DisplayLayout.Bands(0).Columns("Child_Contact").Hidden = True
            ugContacts.DisplayLayout.Bands(0).Columns("ContactAssocID").Hidden = True
            ugContacts.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
            ugContacts.DisplayLayout.Bands(0).Columns("ModuleID").Hidden = True
            ugContacts.DisplayLayout.Bands(0).Columns("ContactType").Hidden = True

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnModifyContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifyContact.Click
        Try
            '   If ugContacts.Rows.Count <= 0 Then Exit Sub

            If ugContacts.ActiveRow Is Nothing Then
                MsgBox("Select row to Modify.")
                Exit Sub
            End If
            Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow = ugContacts.ActiveRow
            If IsNothing(ContactFrm) Then
                ContactFrm = New Contacts(Integer.Parse(dr.Cells("EntityID").Value), CInt(dr.Cells("EntityType").Value), dr.Cells("Module").Value, CInt(dr.Cells("ContactID").Value), dr, pConStruct, "MODIFY")
                AddHandler ContactFrm.Closing, AddressOf frmContactsClosing
                AddHandler ContactFrm.Closed, AddressOf frmContactsClosed
            End If

            ContactFrm.ShowDialog()


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub frmContactsClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Try

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub frmContactsClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not ContactFrm Is Nothing Then
            ContactFrm = Nothing
        End If
    End Sub

    Private Sub btnDeleteContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteContact.Click
        Try
            If ugContacts.Rows.Count <= 0 Then Exit Sub

            If ugContacts.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete.")
                Exit Sub
            End If
            If (CInt(ugContacts.ActiveRow.Cells("EntityID").Value) <> pConStruct.entityID) Or (CInt(ugContacts.ActiveRow.Cells("ModuleID").Value) <> pConStruct.ModuleID) Then
                MsgBox("Selected contact is not associated with the current entity and cannot be deleted.")
                Exit Sub
            End If
            result = MessageBox.Show("Are you Sure you want to Delete this Record?", "Contact", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.No Then Exit Sub

            pConStruct.Remove(ugContacts.ActiveRow.Cells("EntityAssocID").Text, CType(UIUtilsGen.ModuleID.ContactManagement, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            ugContacts.ActiveRow.Delete(False)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub btnAssociateContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAssociateContact.Click
        Try
            If ugContacts.Rows.Count <= 0 Then Exit Sub

            If ugContacts.ActiveRow Is Nothing Then
                MsgBox("Select row to Associate.")
                Exit Sub
            End If
            Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow = ugContacts.ActiveRow
            Dim objContact As Contacts
            If ((CInt(ugContacts.ActiveRow.Cells("EntityID").Value) = pConStruct.entityID) And (CInt(ugContacts.ActiveRow.Cells("ModuleID").Value) = pConStruct.ModuleID)) Then
                MsgBox("Selected contact is already associated with the current entity")
                Exit Sub
            End If
            If IsNothing(ContactFrm) Then
                ContactFrm = New Contacts(Integer.Parse(dr.Cells("EntityID").Value), CInt(dr.Cells("EntityType").Value), dr.Cells("Module").Value, CInt(dr.Cells("ContactID").Value), dr, pConStruct, "ASSOCIATE")
                AddHandler ContactFrm.Closing, AddressOf frmContactsClosing
                AddHandler ContactFrm.Closed, AddressOf frmContactsClosed
            End If

            ContactFrm.ShowDialog()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub chkShowRelatedContacts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowRelatedContacts.CheckedChanged
        Try
            If chkShowRelatedContacts.Checked And Not dsContacts Is Nothing Then
                dsContacts.Tables(0).DefaultView.RowFilter = ""
                ugContacts.DataSource = dsContacts.Tables(0).DefaultView
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub chkShowContactsforAllModules_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowContactsforAllModules.CheckedChanged
        Try
            If chkShowContactsforAllModules.Checked And Not dsContacts Is Nothing Then
                dsContacts.Tables(0).DefaultView.RowFilter = ""
                ugContacts.DataSource = dsContacts.Tables(0).DefaultView
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkShowActiveOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowActiveOnly.CheckedChanged
        Try
            If chkShowActiveOnly.Checked And Not dsContacts Is Nothing Then
                dsContacts.Tables(0).DefaultView.RowFilter = "ACTIVE = 1"
                ugContacts.DataSource = dsContacts.Tables(0).DefaultView
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
